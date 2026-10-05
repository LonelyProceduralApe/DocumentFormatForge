"""docx 底层读写与 OOXML 辅助函数。

基于 python-docx，并对「字符级缩进」「固定行距」「东亚字体」等 python-docx
未直接暴露的属性做原始 XML 操作，确保 WPS/Word 正确识别。
"""
from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt

MM_TO_TWIPS = 56.7

_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def set_run_font(run, east_asia: str, ascii_font: str, size_pt: float) -> None:
    """设置 run 的东亚/西文字体与字号（含复杂文种）。"""
    run.font.name = ascii_font
    run.font.size = Pt(size_pt)
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), east_asia)
    rFonts.set(qn("w:cs"), ascii_font)
    sz = rPr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        rPr.append(sz)
    sz.set(qn("w:val"), str(int(round(size_pt * 2))))
    szCs = rPr.find(qn("w:szCs"))
    if szCs is None:
        szCs = OxmlElement("w:szCs")
        rPr.append(szCs)
    szCs.set(qn("w:val"), str(int(round(size_pt * 2))))


def set_line_exact(paragraph, pt: float) -> None:
    """固定行距（exact），段前段后置 0。"""
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        pPr.append(spacing)
    spacing.set(qn("w:line"), str(int(round(pt * 20))))
    spacing.set(qn("w:lineRule"), "exact")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")


def set_indent_chars(paragraph, size_pt: float, first_line_chars=None,
                     left_chars=None, right_chars=None) -> None:
    """字符级缩进（Word/WPS 按字宽计算），并写入磅值兜底。"""
    pPr = paragraph._p.get_or_add_pPr()
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    per_char_twips = size_pt * 20  # 每字按当前字号换算
    if first_line_chars is not None:
        ind.set(qn("w:firstLineChars"), str(int(round(first_line_chars * 100))))
        ind.set(qn("w:firstLine"), str(int(round(first_line_chars * per_char_twips))))
    if left_chars is not None:
        ind.set(qn("w:leftChars"), str(int(round(left_chars * 100))))
        ind.set(qn("w:left"), str(int(round(left_chars * per_char_twips))))
    if right_chars is not None:
        ind.set(qn("w:rightChars"), str(int(round(right_chars * 100))))
        ind.set(qn("w:right"), str(int(round(right_chars * per_char_twips))))


def set_alignment(paragraph, align: str) -> None:
    paragraph.alignment = _ALIGN_MAP.get(align, WD_ALIGN_PARAGRAPH.LEFT)


def clear_paragraph_runs(paragraph) -> None:
    """清空段落内的所有 run，保留段落对象。"""
    for r in list(paragraph._p.findall(qn("w:r"))):
        paragraph._p.remove(r)


def apply_page_setup(section, rules: dict) -> None:
    """页面设置：页边距（纵向/横向都套用）；纵向节另设纸张尺寸与文档网格。"""
    pg = rules["page"]
    is_landscape = section.page_width > section.page_height
    # 页边距与页眉页脚（横向节也套用公文页边距）
    section.top_margin = Mm(pg["margin_top_mm"])
    section.bottom_margin = Mm(pg["margin_bottom_mm"])
    section.left_margin = Mm(pg["margin_left_mm"])
    section.right_margin = Mm(pg["margin_right_mm"])
    section.gutter = Mm(pg["gutter_mm"])
    section.header_distance = Mm(pg["header_mm"])
    section.footer_distance = Mm(pg["footer_mm"])
    if is_landscape:
        return  # 横向节：保留方向与纸张尺寸，仅改页边距
    section.page_width = Mm(pg["width_mm"])
    section.page_height = Mm(pg["height_mm"])

    # 文档网格：指定行网格（每面 22 行 → 行距 29 磅 ≈ 580 twips）
    sectPr = section._sectPr
    docGrid = sectPr.find(qn("w:docGrid"))
    if docGrid is None:
        docGrid = OxmlElement("w:docGrid")
        sectPr.append(docGrid)
    line_pitch = int(round(rules["line_spacing_pt"].get("body", 29) * 20))
    docGrid.set(qn("w:type"), "lines")
    docGrid.set(qn("w:linePitch"), str(line_pitch))
    docGrid.set(qn("w:charSpace"), "0")


def iter_paragraphs(document):
    """遍历正文段落，返回 [(index, paragraph, text)]，含空段。"""
    result = []
    for idx, p in enumerate(document.paragraphs):
        text = p.text or ""
        result.append((idx, p, text))
    return result


def read_paragraphs(document):
    """读取段落文本与图形标记（供分类器使用）。"""
    result = []
    for i, p in enumerate(document.paragraphs):
        has_pict = bool(p._p.findall(".//" + qn("w:pict"))) or \
            bool(p._p.findall(".//" + qn("w:drawing")))
        result.append({"idx": i, "text": (p.text or "").strip(), "has_pict": has_pict})
    return result
