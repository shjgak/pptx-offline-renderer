#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Lightweight PPTX -> PNG renderer (offline, Pillow + system CJK font).

Approximate rendering intended for *visual QA*: it reconstructs slide layout,
colors, shapes, embedded pictures, and text so an agent (or a human) can "see"
a .pptx without PowerPoint / LibreOffice installed. It is NOT pixel-perfect and
does not reproduce exact typography — use it to catch overflow, out-of-bounds
shapes, missing pictures, and gross layout problems.

Model-agnostic: the renderer only produces PNG files. Any vision-capable model
(e.g. DeepSeek V4 Flash / Pro, Qwen-VL, GLM-4V, Step-1V, Claude, GPT-4V ...) can
consume them by attaching slide_N.png or contact_sheet.png as an image input.

Usage:
    python render_pptx.py <file.pptx> [output_dir] [--width N] [--no-contact]

Options:
    --width N      output width in px (default 2000). Lower (e.g. 1280) for
                   lightweight vision models with a small image budget (Flash
                   class); higher (e.g. 2400) for strong vision models (Pro class).
    --no-contact   skip generating the combined contact_sheet.png.

Outputs one PNG per slide (slide_1.png ...) plus a contact_sheet.png (unless
disabled), and prints any OVERFLOW / BOUNDS / PIC issues found.
"""
import sys, os, zipfile, posixpath, argparse
import xml.etree.ElementTree as ET
from PIL import Image, ImageDraw, ImageFont

Z = None  # open zipfile handle (global, for pic media)

EMU_PER_IN = 914400
PX_W = 2000  # output width in px

# ---------------------------------------------------------------------------
# CJK font auto-detection (cross-platform)
# ---------------------------------------------------------------------------
def detect_cjk_font():
    """Return (path, index) for a usable CJK font, or (None, 0)."""
    candidates = [
        # macOS
        ("/System/Library/Fonts/STHeiti Medium.ttc", 0),
        ("/System/Library/Fonts/PingFang.ttc", 0),
        ("/Library/Fonts/Arial Unicode.ttf", 0),
        # Windows
        ("C:/Windows/Fonts/msyh.ttc", 1),   # Microsoft YaHei
        ("C:/Windows/Fonts/simhei.ttf", 0),  # SimHei
        ("C:/Windows/Fonts/simsun.ttc", 1),  # SimSun
        # Linux (Noto / WenQuanYi)
        ("/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc", 0),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
        ("/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc", 0),
        ("/usr/share/fonts/truetype/wqy/wqy-microhei.ttc", 0),
    ]
    for path, idx in candidates:
        if os.path.exists(path):
            return path, idx
    return None, 0

FONT_PATH, FONT_INDEX = detect_cjk_font()


def lname(tag):
    return tag.split('}')[-1] if '}' in tag else tag


def find(el, name, ns=None):
    for c in el.iter():
        if lname(c.tag) == name:
            return c
    return None


def findall(el, name):
    return [c for c in el.iter() if lname(c.tag) == name]


class Theme:
    def __init__(self, xml_bytes):
        self.scheme = {}
        try:
            root = ET.fromstring(xml_bytes)
        except Exception:
            return
        for cs in findall(root, 'clrScheme'):
            for child in cs:
                k = lname(child.tag)  # dk1 lt1 dk2 lt2 accent1..6
                srgb = find(child, 'srgbClr')
                sysclr = find(child, 'sysClr')
                if srgb is not None:
                    self.scheme[k] = '#' + srgb.get('val')
                elif sysclr is not None:
                    v = sysclr.get('lastClr') or sysclr.get('val')
                    if v and v.lower() not in ('window', 'windowtext'):
                        self.scheme[k] = '#' + v
                    else:
                        self.scheme[k] = '#FFFFFF' if v == 'window' else '#000000'
        # fallbacks
        self.scheme.setdefault('dk1', '#000000')
        self.scheme.setdefault('lt1', '#FFFFFF')
        self.scheme.setdefault('dk2', '#000000')
        self.scheme.setdefault('lt2', '#FFFFFF')


def resolve_color(el, theme):
    """el is an <a:solidFill> or scheme/srgb child. Return (r,g,b) or None."""
    if el is None:
        return None
    srgb = find(el, 'srgbClr')
    if srgb is not None:
        h = srgb.get('val', '000000')
        return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))
    sch = find(el, 'schemeClr')
    if sch is not None:
        key = sch.get('val')
        h = theme.scheme.get(key, '#000000')
        return tuple(int(h[i:i + 2], 16) for i in (1, 3, 5))
    sysclr = find(el, 'sysClr')
    if sysclr is not None:
        v = sysclr.get('lastClr') or sysclr.get('val')
        if v and v.lower() not in ('window', 'windowtext'):
            try:
                return tuple(int(v[i:i + 2], 16) for i in (0, 2, 4))
            except Exception:
                pass
        return (255, 255, 255) if v == 'window' else (0, 0, 0)
    return None


def parse_fill(spPr, theme):
    sf = find(spPr, 'solidFill')
    if sf is not None:
        return resolve_color(find(sf, 'srgbClr') or find(sf, 'schemeClr') or sf, theme)
    if find(spPr, 'noFill') is not None:
        return None
    return None


def parse_line(spPr, theme):
    ln = find(spPr, 'ln')
    if ln is None:
        return None, 0
    sf = find(ln, 'solidFill')
    color = resolve_color(find(sf, 'srgbClr') or find(sf, 'schemeClr') or sf, theme) if sf is not None else (0, 0, 0)
    w = ln.get('w')
    wpx = 1
    if w:
        try:
            wpx = max(1, int(int(w) * PX_W / EMU_PER_IN / 72))
        except Exception:
            wpx = 1
    return color, wpx


def get_text_runs(txBody):
    """Return list of paragraphs; each paragraph = list of runs.
    run = (text, size_pt, bold, italic, color_rgb, align). align from pPr."""
    paras = []
    for p in findall(txBody, 'p'):
        pPr = find(p, 'pPr')
        algn = pPr.get('algn', 'l') if pPr is not None else 'l'
        runs = []
        for r in findall(p, 'r'):
            rPr = find(r, 'rPr')
            t = find(r, 't')
            txt = t.text if t is not None and t.text else ''
            sz = 18.0
            bold = False
            italic = False
            color = (0, 0, 0)
            if rPr is not None:
                if rPr.get('sz'):
                    try:
                        sz = int(rPr.get('sz')) / 100.0
                    except Exception:
                        sz = 18.0
                bold = rPr.get('b') == '1'
                italic = rPr.get('i') == '1'
                sf = find(rPr, 'solidFill')
                if sf is not None:
                    c = resolve_color(find(sf, 'srgbClr') or find(sf, 'schemeClr') or sf, global_theme)
                    if c:
                        color = c
            runs.append((txt, sz, bold, italic, color))
        paras.append((runs, algn))
    return paras


def _is_cjk(ch):
    c = ord(ch)
    return (0x4E00 <= c <= 0x9FFF or 0x3400 <= c <= 0x4DBF or
            0x3040 <= c <= 0x309F or 0x30A0 <= c <= 0x30FF or
            0xAC00 <= c <= 0xD7AF or 0xFF00 <= c <= 0xFFEF)


def wrap_text(draw, runs, max_w, font_factory):
    """Yield list of lines; each line = list of (seg_text, font, color, bold)."""
    tokens = []
    for (txt, sz, bold, italic, color) in runs:
        font = font_factory(sz, bold)
        i = 0
        while i < len(txt):
            ch = txt[i]
            if ch in '\r\n':
                tokens.append(('BR', None, None, None))
                i += 1
            elif ch == ' ':
                tokens.append((' ', font, color, bold))
                i += 1
            elif _is_cjk(ch):
                tokens.append((ch, font, color, bold))
                i += 1
            else:
                j = i
                while j < len(txt) and not _is_cjk(txt[j]) and txt[j] not in ' \r\n':
                    j += 1
                tokens.append((txt[i:j], font, color, bold))
                i = j

    lines = []
    cur = []
    cur_w = 0
    for tok in tokens:
        if tok[0] == 'BR':
            lines.append(cur)
            cur = []
            cur_w = 0
            continue
        text, font, color, bold = tok
        w = draw.textlength(text, font=font)
        if cur_w + w > max_w and cur:
            lines.append(cur)
            cur = []
            cur_w = 0
        cur.append(tok)
        cur_w += w
    if cur:
        lines.append(cur)
    return lines


def render_slide(slide_xml, theme, scale, out_path, rels_map):
    root = ET.fromstring(slide_xml)
    W = int(PX_W)
    H = int(PX_W * slide_h_emu / slide_w_emu)
    img = Image.new('RGB', (W, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    sp_tree = find(root, 'spTree')
    if sp_tree is None:
        sp_tree = root
    shapes = [c for c in sp_tree if lname(c.tag) in ('sp', 'pic', 'graphicFrame')]
    issues = []
    for sp in shapes:
        spPr = find(sp, 'spPr')
        if spPr is None:
            continue
        xfrm = find(spPr, 'xfrm')
        if xfrm is None:
            continue
        off = find(xfrm, 'off')
        ext = find(xfrm, 'ext')
        if off is None or ext is None:
            continue
        x = int(int(off.get('x')) * scale)
        y = int(int(off.get('y')) * scale)
        w = int(int(ext.get('cx')) * scale)
        h = int(int(ext.get('cy')) * scale)
        if y + h > H + 2 or x + w > W + 2:
            issues.append(('BOUNDS', f"y+h={y + h}px(>{H}) x+w={x + w}px(>{W})"))
        # picture: composite embedded media
        if lname(sp.tag) == 'pic':
            blip = find(sp, 'blip')
            rid = None
            if blip is not None:
                for k, v in blip.attrib.items():
                    if k.endswith('}embed'):
                        rid = v
            target = rels_map.get(rid) if rid else None
            if target:
                try:
                    pim = Image.open(Z.open(target)).convert('RGB')
                    pim = pim.resize((max(1, w), max(1, h)))
                    img.paste(pim, (x, y))
                except Exception as e:
                    issues.append(('PIC', str(e)))
            continue
        fill = parse_fill(spPr, theme)
        line_color, line_w = parse_line(spPr, theme)
        if fill is not None:
            d.rectangle([x, y, x + w, y + h], fill=fill, outline=line_color, width=line_w)
        else:
            if line_color is not None:
                d.rectangle([x, y, x + w, y + h], outline=line_color, width=line_w)
        txBody = find(sp, 'txBody')
        if txBody is not None:
            bodyPr = find(txBody, 'bodyPr')
            anchor = bodyPr.get('anchor', 't') if bodyPr is not None else 't'
            l_ins = t_ins = int(0.05 * EMU_PER_IN * scale)
            if bodyPr is not None:
                li = bodyPr.get('lIns')
                ti = bodyPr.get('tIns')
                ri = bodyPr.get('rIns')
                bi = bodyPr.get('bIns')
                if li:
                    l_ins = max(l_ins, int(int(li) * scale))
                if ti:
                    t_ins = max(t_ins, int(int(ti) * scale))
                if ri:
                    l_ins = max(l_ins, int(int(ri) * scale))
                if bi:
                    t_ins = max(t_ins, int(int(bi) * scale))
            paras = get_text_runs(txBody)
            all_lines = []
            for (runs, algn) in paras:
                if not runs or not any(r[0].strip() for r in runs):
                    continue
                wrapped = wrap_text(d, runs, w - 2 * l_ins, font_factory)
                for ln in wrapped:
                    all_lines.append((ln, algn))
            total_h = 0
            line_heights = []
            for (ln, algn) in all_lines:
                if not ln:
                    lh = int(0.28 * EMU_PER_IN * scale)
                    line_heights.append(lh)
                    total_h += lh
                    continue
                maxsz = max((seg[1].size for seg in ln), default=18)
                lh = int(maxsz * 1.25) + 2
                line_heights.append(lh)
                total_h += lh
            if total_h > h - 2:
                issues.append(('OVERFLOW', f"text_h={total_h}px box_h={h}px ({(total_h - h):+d})"))
            if anchor == 'ctr':
                cy = y + (h - total_h) // 2
            elif anchor == 'b':
                cy = y + (h - total_h) - t_ins
            else:
                cy = y + t_ins
            idx = 0
            for (ln, algn) in all_lines:
                lh = line_heights[idx]
                idx += 1
                if not ln:
                    cy += lh
                    continue
                lw = sum(d.textlength(seg[0], font=seg[1]) for seg in ln)
                if algn == 'ctr':
                    lx = x + (w - lw) // 2
                elif algn == 'r':
                    lx = x + w - l_ins - lw
                else:
                    lx = x + l_ins
                ty = cy
                for (seg_text, font, color, bold) in ln:
                    d.text((lx, ty), seg_text, font=font, fill=color)
                    lx += d.textlength(seg_text, font=font)
                cy += lh
    img.save(out_path)
    return out_path, issues


def font_factory(sz_pt, bold):
    px = max(6, int(sz_pt * PX_W / (slide_w_emu / EMU_PER_IN) / 72))
    try:
        if FONT_PATH:
            f = ImageFont.truetype(FONT_PATH, px, index=FONT_INDEX)
        else:
            f = ImageFont.load_default()
    except Exception:
        f = ImageFont.load_default()
    return f


global_theme = None
slide_w_emu = 12192000
slide_h_emu = 6858000


def main():
    global global_theme, slide_w_emu, slide_h_emu, Z, PX_W
    ap = argparse.ArgumentParser(
        description="Offline PPTX -> PNG renderer (no PowerPoint/LibreOffice needed).")
    ap.add_argument("pptx", help="path to the .pptx file")
    ap.add_argument("outdir", nargs="?", default=None, help="output directory (default: next to the pptx)")
    ap.add_argument("--width", type=int, default=2000,
                    help="output width in px (default 2000). Lower for lightweight vision "
                         "models (e.g. 1280 for Flash class), higher for strong ones (e.g. 2400 for Pro class).")
    ap.add_argument("--no-contact", action="store_true", help="do not generate contact_sheet.png")
    args = ap.parse_args()

    PX_W = max(400, int(args.width))
    pptx = args.pptx
    outdir = args.outdir or os.path.dirname(os.path.abspath(pptx))
    os.makedirs(outdir, exist_ok=True)
    if not FONT_PATH:
        print("[警告] 未找到中文字体，中文将显示为方框；请先安装任意 CJK 字体后重试。")
        print("WARNING: no CJK font found; CJK glyphs will render as boxes. Install a CJK font.")
    print(f"[渲染] {pptx}  ->  {outdir}  (width={PX_W}px)")
    print(f"[render] {pptx} -> {outdir} (width={PX_W}px)")
    z = zipfile.ZipFile(pptx)
    Z = z
    names = z.namelist()
    # theme
    theme_path = next((n for n in names if n.startswith('ppt/theme/') and n.endswith('.xml')), None)
    theme = Theme(z.read(theme_path)) if theme_path else Theme(b'')
    global_theme = theme
    # slide size
    pres = ET.fromstring(z.read('ppt/presentation.xml'))
    sldSz = find(pres, 'sldSz')
    if sldSz is not None:
        slide_w_emu = int(sldSz.get('cx'))
        slide_h_emu = int(sldSz.get('cy'))
    scale = PX_W / slide_w_emu
    # slides
    slide_files = sorted(
        [n for n in names if n.startswith('ppt/slides/slide') and n.endswith('.xml')],
        key=lambda s: int(''.join(ch for ch in s if ch.isdigit())))
    outs = []
    for sf in slide_files:
        idx = ''.join(ch for ch in sf if ch.isdigit())
        rels_name = 'ppt/slides/_rels/slide%s.xml.rels' % idx
        rels_map = {}
        if rels_name in names:
            rroot = ET.fromstring(z.read(rels_name))
            for rel in rroot:
                rid = rel.get('Id')
                tgt = rel.get('Target')
                if rid and tgt and not tgt.startswith('http'):
                    full = posixpath.normpath('ppt/slides/' + tgt)
                    rels_map[rid] = full
        out = os.path.join(outdir, f'slide_{idx}.png')
        _, issues = render_slide(z.read(sf), theme, scale, out, rels_map)
        outs.append(out)
        print(f"[完成] 第{idx}页 -> {out}")
        print(f"rendered {out}")
        for it in issues:
            print(f"   !!! {it[0]} {it[1]}")
    # contact sheet
    if outs and not args.no_contact:
        imgs = [Image.open(o) for o in outs]
        cols = len(imgs)
        cw, ch = imgs[0].size
        pad = 20
        sheet = Image.new('RGB', (cols * cw + (cols + 1) * pad, ch + 2 * pad), (230, 230, 230))
        for i, im in enumerate(imgs):
            sheet.paste(im, (pad + i * (cw + pad), pad))
        sheet_path = os.path.join(outdir, 'contact_sheet.png')
        sheet.save(sheet_path)
        print(f"[完成] 拼图 -> {sheet_path}")
        print(f"contact {sheet_path}")


if __name__ == '__main__':
    main()
