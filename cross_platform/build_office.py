#!/usr/bin/env python3
"""跨平台 Office 生成引擎（macOS / Linux / Windows 通用）。

不依赖任何 COM / 闭源软件，用 python-pptx / python-docx / openpyxl 生成真实
docx / xlsx / pptx，再用本机 LibreOffice (soffice) 一键导出 PDF。

子命令：
    python build_office.py ppt   --json spec.json --out deck.pptx [--pdf]
    python build_office.py docx  --json spec.json --out report.docx [--pdf]
    python build_office.py xlsx  --json spec.json --out data.xlsx [--pdf]

中文字体：在 macOS 上自动把 SimHei/微软雅黑 映射为 PingFang SC / Songti SC，
确保导出的 PDF 中文不乱码。
"""
import os
import sys
import json
import shutil
import argparse
import subprocess

# ── 依赖（用户需 pip install python-pptx python-docx openpyxl）─────────────
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from docx import Document
from docx.shared import Pt as DocPt, Inches, RGBColor as DocRGB
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ── 中文字体映射（macOS 无 SimHei/微软雅黑，用系统自带替代）─────────────────
FONT_MAP = {
    "SimHei": "PingFang SC",
    "微软雅黑": "PingFang SC",
    "Microsoft YaHei": "PingFang SC",
    "SimSun": "Songti SC",
    "宋体": "Songti SC",
    "KaiTi": "Kaiti SC",
    "黑体": "PingFang SC",
}
DEFAULT_FONT = "PingFang SC"

# 逻辑画布（与官方 harness 一致，960×540）
SCALE = 12700  # EMU per logical px
CANVAS_W, CANVAS_H = 960, 540


def cmap(font):
    return FONT_MAP.get(font, font) or DEFAULT_FONT


def hex2rgb(h):
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def px(v):
    return Emu(int(float(v) * SCALE))


# ── soffice 定位与 PDF 导出 ────────────────────────────────────────────────
def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    mac = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
    if os.path.isfile(mac):
        return mac
    return None


def export_pdf(src_path, timeout=120):
    lo = find_soffice()
    if not lo:
        raise RuntimeError("未找到 LibreOffice(soffice)。Mac: brew install --cask libreoffice")
    out_dir = os.path.dirname(os.path.abspath(src_path))
    with subprocess.Popen(
        [lo, "--headless", "--nologo", "--nofirststartwizard",
         "--convert-to", "pdf", "--outdir", out_dir, os.path.abspath(src_path)],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    ) as proc:
        proc.wait(timeout=timeout)
    base = os.path.splitext(os.path.basename(src_path))[0]
    pdf = os.path.join(out_dir, base + ".pdf")
    if not os.path.exists(pdf):
        raise RuntimeError("PDF 导出失败，未生成 " + pdf)
    return pdf


# ── PPTX ────────────────────────────────────────────────────────────────────
def _set_ea(run, font):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {"typeface": font})
        rPr.append(ea)
    else:
        ea.set("typeface", font)


def _style_run(run, spec):
    font = cmap(spec.get("font", DEFAULT_FONT))
    if "fs" in spec:
        run.font.size = Pt(float(spec["fs"]))
    run.font.bold = bool(spec.get("bold", False))
    run.font.italic = bool(spec.get("italic", False))
    if "color" in spec:
        run.font.color.rgb = hex2rgb(spec["color"])
    run.font.name = font
    _set_ea(run, font)


_ALIGN_MAP = {1: PP_ALIGN.LEFT, 2: PP_ALIGN.CENTER, 3: PP_ALIGN.RIGHT,
                "1": PP_ALIGN.LEFT, "2": PP_ALIGN.CENTER, "3": PP_ALIGN.RIGHT,
                "left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}


def _resolve_align(el):
    a = el.get("align", "left")
    return _ALIGN_MAP.get(a, _ALIGN_MAP.get(str(a).lower(), PP_ALIGN.LEFT))


def _add_text(slide, el):
    tb = slide.shapes.add_textbox(px(el["x"]), px(el["y"]), px(el["w"]), px(el["h"]))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    al = _resolve_align(el)
    lines = el["text"].split("\n") if isinstance(el.get("text"), str) else el.get("text", [])
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = al
        if el.get("line_spacing"):
            p.line_spacing = float(el["line_spacing"])
        r = p.add_run(); r.text = line
        _style_run(r, el)
    return tb


def _add_line(slide, el):
    color = hex2rgb(el.get("color", "#000000"))
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, px(el["x"]), px(el["y"]),
                                   px(el["w"]), px(el["h"]))
    shape.fill.solid(); shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False
    return shape


def _card_color(card, el):
    return hex2rgb(card.get("color", el.get("title_color", "#660874")))


def _add_grid_card(slide, x, y, w, h, card, el):
    color = _card_color(card, el)
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(x), px(y), px(w), px(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shape.line.color.rgb = color; shape.line.width = Pt(1)
    shape.shadow.inherit = False
    tf = shape.text_frame; tf.word_wrap = True
    tf.margin_left = Pt(10); tf.margin_right = Pt(10)
    tf.margin_top = Pt(8); tf.margin_bottom = Pt(8)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    font = el.get("font", DEFAULT_FONT)
    gray = RGBColor(0x44, 0x44, 0x44)
    if card.get("icon"):
        p = tf.paragraphs[0]; r = p.add_run(); r.text = card["icon"]
        r.font.bold = True; r.font.size = Pt(13); r.font.color.rgb = color
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    if card.get("num") is not None:
        p = tf.paragraphs[0] if not card.get("icon") else tf.add_paragraph()
        r = p.add_run(); r.text = str(card["num"])
        r.font.bold = True; r.font.size = Pt(el.get("num_fs", 26)); r.font.color.rgb = color
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    if card.get("title"):
        p = tf.add_paragraph()
        r = p.add_run(); r.text = card["title"]
        r.font.bold = True; r.font.size = Pt(el.get("title_fs", 15)); r.font.color.rgb = color
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    label = card.get("label")
    body = card.get("desc") or card.get("body")
    if label:
        p = tf.add_paragraph(); r = p.add_run(); r.text = label
        r.font.size = Pt(el.get("label_fs", 12)); r.font.color.rgb = gray
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    if body:
        p = tf.add_paragraph(); r = p.add_run(); r.text = body
        r.font.size = Pt(el.get("body_fs", 12)); r.font.color.rgb = gray
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    if card.get("sub"):
        p = tf.add_paragraph(); r = p.add_run(); r.text = card["sub"]
        r.font.size = Pt(el.get("sub_fs", 12)); r.font.color.rgb = gray
        r.font.name = cmap(font); _set_ea(r, cmap(font))


def _add_wide_card(slide, x, y, w, h, card, el):
    color = _card_color(card, el)
    sep = hex2rgb(el.get("sep_color", "#CCCCCC"))
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(x), px(y), px(w), px(h))
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    shape.line.color.rgb = sep; shape.line.width = Pt(0.75)
    shape.shadow.inherit = False
    tf = shape.text_frame; tf.word_wrap = True
    tf.margin_left = Pt(14); tf.margin_top = Pt(6); tf.margin_bottom = Pt(6)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    font = el.get("font", DEFAULT_FONT)
    gray = RGBColor(0x55, 0x55, 0x55)
    p = tf.paragraphs[0]
    if card.get("num") is not None:
        r = p.add_run(); r.text = str(card["num"]) + "  "
        r.font.bold = True; r.font.size = Pt(18); r.font.color.rgb = color
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    if card.get("title"):
        r = p.add_run(); r.text = card["title"]
        r.font.bold = True; r.font.size = Pt(el.get("title_fs", 15)); r.font.color.rgb = hex2rgb(el.get("title_color", "#1A1A1A"))
        r.font.name = cmap(font); _set_ea(r, cmap(font))
    if card.get("sub"):
        p2 = tf.add_paragraph(); r = p2.add_run(); r.text = card["sub"]
        r.font.size = Pt(el.get("sub_fs", 12)); r.font.color.rgb = gray
        r.font.name = cmap(font); _set_ea(r, cmap(font))


def _add_bullets(slide, el):
    tb = slide.shapes.add_textbox(px(el["x"]), px(el["y"]), px(el["w"]), px(el["h"]))
    tf = tb.text_frame; tf.word_wrap = True
    items = el.get("items", [])
    ordered = el.get("ordered", False)
    for i, it in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(6)
        prefix = (f"{i+1}. " if ordered else "• ")
        r = p.add_run(); r.text = prefix + it
        _style_run(r, el)
    return tb


def _add_table(slide, el):
    raw = el.get("rows")
    # 兼容两种 schema：
    #  旧版：rows 为二维数据数组；新版：rows=行数(int), cols=列数(int), data=二维数组
    if isinstance(raw, list) and raw and isinstance(raw[0], list):
        data = raw
        header = el.get("header", False)
    else:
        data = el.get("data", [])
        header = el.get("header", True)
    if not data:
        return None
    nrows = len(data)
    ncols = max((len(r) for r in data if isinstance(r, list)), default=0)
    if ncols == 0:
        return None
    gtbl = slide.shapes.add_table(nrows, ncols, px(el["x"]), px(el["y"]), px(el["w"]), px(el["h"]))
    table = gtbl.table
    th_fs = float(el.get("th_fs", el.get("fs", 14)))
    td_fs = float(el.get("td_fs", el.get("fs", 14)))
    hcolor = hex2rgb(el.get("header_color", "#006B3F"))
    for ri, row in enumerate(data):
        for ci in range(ncols):
            val = row[ci] if isinstance(row, list) and ci < len(row) else ""
            cell = table.cell(ri, ci)
            cell.text = "" if val is None else str(val)
            para = cell.text_frame.paragraphs[0]
            run = para.runs[0] if para.runs else para.add_run()
            if run.text == "":
                run.text = "" if val is None else str(val)
            run.font.size = Pt(th_fs if (header and ri == 0) else td_fs)
            run.font.name = cmap(el.get("font", DEFAULT_FONT))
            _set_ea(run, cmap(el.get("font", DEFAULT_FONT)))
            if header and ri == 0:
                run.font.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                cell.fill.solid(); cell.fill.fore_color.rgb = hcolor
            else:
                run.font.color.rgb = hex2rgb(el.get("td_color", "#333333"))
                cell.fill.solid(); cell.fill.fore_color.rgb = hex2rgb(el.get("cell_color", "#FFFFFF"))
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    return gtbl


def _add_cards(slide, el, cols=None, rows=None, wide=False):
    items = el.get("cards") or el.get("items", [])
    if wide:
        x = el.get("x", 60); w = el.get("w", 840)
        y = el.get("start_y", el.get("y", 90))
        ih = el.get("item_h", 48); gap = el.get("gap", 2)
        for idx, card in enumerate(items):
            cy = y + idx * (ih + gap)
            _add_wide_card(slide, x, cy, w, ih, card, el)
        return
    if cols is None:
        cols = el.get("cols", 3)
    if rows is None:
        rows = max(1, (len(items) + cols - 1) // cols)
    x = el.get("x", 60)
    y = el.get("start_y", el.get("y", 130))
    w = el.get("w", 840)
    h = el.get("h", 360)
    gap = 12
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    for idx, card in enumerate(items[: cols * rows]):
        r, c = divmod(idx, cols)
        cx = x + c * (cw + gap)
        cy = y + r * (ch + gap)
        _add_grid_card(slide, cx, cy, cw, ch, card, el)


def _add_tagline(slide, el):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, px(el["x"]), px(el["y"]), px(el["w"]), px(el["h"]))
    shape.fill.solid(); shape.fill.fore_color.rgb = hex2rgb(el.get("color", "#006B3F"))
    shape.line.fill.background()
    tf = shape.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = el.get("text", "")
    r.font.size = Pt(float(el.get("fs", 18)))
    r.font.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r.font.name = cmap(el.get("font", DEFAULT_FONT)); _set_ea(r, cmap(el.get("font", DEFAULT_FONT)))


def _add_image_placeholder(slide, el, name):
    x, y = px(el.get("x", 0)), px(el.get("y", 0))
    w, h = px(el.get("w", 200)), px(el.get("h", 150))
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(0xF2, 0xF2, 0xF2)
    shape.line.color.rgb = RGBColor(0xCC, 0xCC, 0xCC); shape.line.width = Pt(1)
    shape.shadow.inherit = False
    tf = shape.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    font = el.get("font", DEFAULT_FONT)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = name
    r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    r.font.name = cmap(font); _set_ea(r, cmap(font))
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r2 = p2.add_run(); r2.text = "（示意图 · 原图未随包提供）"
    r2.font.size = Pt(11); r2.font.color.rgb = RGBColor(0xAA, 0xAA, 0xAA)
    r2.font.name = cmap(font); _set_ea(r2, cmap(font))
    return shape


def _add_image(slide, el):
    src = el.get("src") or el.get("file") or el.get("path") or el.get("image")
    if src and os.path.exists(src):
        return slide.shapes.add_picture(src, px(el["x"]), px(el["y"]), px(el.get("w")), px(el.get("h")))
    # 资源缺失时绘制占位框，避免空白破版
    name = os.path.basename(str(src)) if src else "image"
    return _add_image_placeholder(slide, el, name)


_ELEMENT_HANDLERS = {
    "text": _add_text,
    "bullets": _add_bullets,
    "table": _add_table,
    "image": _add_image,
    "tagline_bar": _add_tagline,
    "line": _add_line,
    "card_list_wide": lambda s, e: _add_cards(s, e, wide=True),
    "cards_2x3": lambda s, e: _add_cards(s, e, 3, 2),
    "cards_2x2_four": lambda s, e: _add_cards(s, e, 2, 2),
    "cards_1x4_info": lambda s, e: _add_cards(s, e, 4, 1),
    "cards_1x3_big": lambda s, e: _add_cards(s, e, 3, 1),
    "card_row_5": lambda s, e: _add_cards(s, e, 5, 1),
}


def build_pptx(spec, out_path, do_pdf, bg_color=None, bg_image=None):
    prs = Presentation()
    prs.slide_width = px(CANVAS_W)
    prs.slide_height = px(CANVAS_H)
    theme = spec.get("theme", {})
    for slide_spec in spec.get("slides", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 背景
        bg = bg_image or slide_spec.get("bg") or spec.get("bg")
        if bg and os.path.exists(bg):
            slide.shapes.add_picture(bg, 0, 0, px(CANVAS_W), px(CANVAS_H))
        elif bg_color:
            slide.background.fill.solid(); slide.background.fill.fore_color.rgb = hex2rgb(bg_color)
        # 每页标题便捷项
        if slide_spec.get("title"):
            _add_text(slide, {"type": "text", "x": 60, "y": 40, "w": CANVAS_W - 120, "h": 70,
                              "text": slide_spec["title"], "fs": 32, "bold": True,
                              "color": theme.get("primary", "#006B3F"), "align": "left",
                              "font": DEFAULT_FONT})
        for el in slide_spec.get("elements", []):
            h = _ELEMENT_HANDLERS.get(el.get("type"))
            if h:
                h(slide, el)
    prs.save(out_path)
    pdf = export_pdf(out_path) if do_pdf else None
    return out_path, pdf


# ── DOCX ────────────────────────────────────────────────────────────────────
def _docx_set_ea(run, font):
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    rFonts.set(qn("w:eastAsia"), font)


def build_docx(spec, out_path, do_pdf):
    doc = Document()
    # 基础样式：中文字体
    normal = doc.styles["Normal"]
    normal.font.name = DEFAULT_FONT
    normal.font.size = DocPt(12)
    _rpr = normal.element.get_or_add_rPr()
    _rf = _rpr.find(qn("w:rFonts"))
    if _rf is None:
        _rf = OxmlElement("w:rFonts"); _rpr.append(_rf)
    _rf.set(qn("w:eastAsia"), DEFAULT_FONT)

    def add_runs(para, text, font=DEFAULT_FONT, size=12, bold=False, color=None):
        run = para.add_run(text)
        run.font.name = font
        run.font.size = DocPt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = DocRGB(*[int(color.lstrip("#")[i:i+2], 16) for i in (0, 2, 4)])
        _docx_set_ea(run, font)
        return run

    title = spec.get("title")
    if title:
        h = doc.add_heading(title, level=0)
    for sec in spec.get("sections", []):
        t = sec.get("type")
        if t == "heading":
            doc.add_heading(sec.get("text", ""), level=sec.get("level", 1))
        elif t == "paragraph":
            p = doc.add_paragraph()
            add_runs(p, sec.get("text", ""), size=sec.get("size", 12), bold=sec.get("bold", False))
        elif t == "bullets":
            for it in sec.get("items", []):
                p = doc.add_paragraph(style="List Bullet" if not sec.get("ordered") else "List Number")
                add_runs(p, it, size=sec.get("size", 12))
        elif t == "table":
            rows = sec.get("rows", [])
            if not rows:
                continue
            ncols = max(len(r) for r in rows)
            table = doc.add_table(rows=len(rows), cols=ncols)
            table.style = "Light Grid Accent 1"
            for ri, row in enumerate(rows):
                for ci, val in enumerate(row):
                    cell = table.cell(ri, ci)
                    cell.text = "" if val is None else str(val)
                    if sec.get("header") and ri == 0:
                        for r in cell.paragraphs[0].runs:
                            r.font.bold = True
        elif t == "image":
            if sec.get("src") and os.path.exists(sec["src"]):
                doc.add_picture(sec["src"], width=Inches(sec.get("width_in", 5)))
        elif t == "page_break":
            doc.add_page_break()
    doc.save(out_path)
    pdf = export_pdf(out_path) if do_pdf else None
    return out_path, pdf


# ── XLSX ────────────────────────────────────────────────────────────────────
def build_xlsx(spec, out_path, do_pdf):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    header_fill = PatternFill("solid", fgColor="006B3F")
    header_font = Font(bold=True, color="FFFFFF", name=DEFAULT_FONT)
    thin = Side(style="thin", color="DDDDDD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for sheet_spec in spec.get("sheets", []):
        ws = wb.create_sheet(title=sheet_spec.get("name", "Sheet1"))
        columns = sheet_spec.get("columns")
        rows = sheet_spec.get("rows", [])
        start = 1
        if columns:
            for ci, col in enumerate(columns, 1):
                c = ws.cell(row=1, column=ci, value=col)
                c.fill = header_fill; c.font = header_font
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.border = border
            start = 2
        for ri, row in enumerate(rows, start):
            for ci, val in enumerate(row, 1):
                c = ws.cell(row=ri, column=ci, value=val)
                c.font = Font(name=DEFAULT_FONT)
                c.border = border
                c.alignment = Alignment(vertical="center")
        # 列宽自适应（简单）
        for ci in range(1, (len(columns) if columns else (len(rows[0]) if rows else 1)) + 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(ci)].width = sheet_spec.get("col_width", 16)
        if sheet_spec.get("freeze_header", True) and columns:
            ws.freeze_panes = "A2"
    wb.save(out_path)
    pdf = export_pdf(out_path) if do_pdf else None
    return out_path, pdf


# ── CLI ─────────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser(description="跨平台 Office 生成引擎 (ppt/docx/xlsx + PDF)")
    sub = ap.add_subparsers(dest="cmd", required=True)

    for name in ("ppt", "docx", "xlsx"):
        sp = sub.add_parser(name, help=f"生成 {name.upper()}")
        sp.add_argument("--json", required=True, help="规格 JSON 路径")
        sp.add_argument("--out", required=True, help="输出文件路径")
        sp.add_argument("--pdf", action="store_true", help="同时导出 PDF (soffice)")
        sp.add_argument("--bg-color", default=None, help="PPT 背景色 (#RRGGBB)")
        sp.add_argument("--bg", default=None, help="PPT 每页背景图路径 (如清华 template_bg.png)")

    args = ap.parse_args()
    with open(args.json, "r", encoding="utf-8") as f:
        spec = json.load(f)

    if args.cmd == "ppt":
        out, pdf = build_pptx(spec, args.out, args.pdf, args.bg_color, args.bg)
    elif args.cmd == "docx":
        out, pdf = build_docx(spec, args.out, args.pdf)
    elif args.cmd == "xlsx":
        out, pdf = build_xlsx(spec, args.out, args.pdf)

    print(f"生成成功: {out}")
    if pdf:
        print(f"PDF 导出: {pdf}")


if __name__ == "__main__":
    main()
