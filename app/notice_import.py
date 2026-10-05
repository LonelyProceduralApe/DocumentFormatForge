"""从《关于进一步规范公文格式的通知.docx》导入/刷新规则，或与当前配置比对校验。

「两者结合」的第二半：把通知里各示范段落的实际格式与页面设置提取出来，
一键更新 rules，或在处理前比对告警。
"""
from __future__ import annotations

import copy
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from . import rules_config

DEFAULT_NOTICE_NAME = "关于进一步规范公文格式的通知.docx"

# 通知中的示范段落 → 角色（用文本前缀定位）
_SAMPLE_MARKERS = [
    ("标题（", "title"),
    ("×××、×××：（主送机关", "addressee"),
    ("××××××××××（正文", "body"),
    ("一、×××××（一级标题", "h1"),
    ("（一）××××（二级标题", "h2"),
    ("1．××××（三级标题", "h3"),
    ("（1）××××（四级标题", "h4"),
    ("附件：1．××××××（", "attachment_note"),
    ("×年×月×日", "date"),
]

_JC_MAP = {"left": "left", "center": "center", "right": "right", "both": "justify",
           "distribute": "justify"}


def default_notice_path() -> Path:
    # 开发环境：工作区根目录；打包后：exe 同目录
    import sys
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / DEFAULT_NOTICE_NAME
    return Path(__file__).resolve().parent.parent / DEFAULT_NOTICE_NAME


def _para_format(p) -> dict:
    info = {}
    pPr = p._p.pPr
    if pPr is not None:
        jc = pPr.find(qn("w:jc"))
        if jc is not None:
            info["align"] = _JC_MAP.get(jc.get(qn("w:val")), "left")
        spacing = pPr.find(qn("w:spacing"))
        if spacing is not None and spacing.get(qn("w:lineRule")) == "exact":
            line = spacing.get(qn("w:line"))
            if line:
                info["line_pt"] = round(int(line) / 20, 2)
    if p.runs:
        rPr = p.runs[0]._r.rPr
        if rPr is not None:
            rFonts = rPr.find(qn("w:rFonts"))
            if rFonts is not None:
                ea = rFonts.get(qn("w:eastAsia"))
                if ea:
                    info["font"] = ea
            sz = rPr.find(qn("w:sz"))
            if sz is not None:
                info["size_pt"] = round(int(sz.get(qn("w:val"))) / 2, 2)
    return info


def _page_from_notice(document) -> dict:
    pg = {}
    sectPr = document.sections[0]._sectPr
    pgSz = sectPr.find(qn("w:pgSz"))
    if pgSz is not None:
        pg["width_mm"] = round(int(pgSz.get(qn("w:w"))) / 56.7, 1)
        pg["height_mm"] = round(int(pgSz.get(qn("w:h"))) / 56.7, 1)
    pgMar = sectPr.find(qn("w:pgMar"))
    if pgMar is not None:
        for key in ("top", "bottom", "left", "right"):
            val = pgMar.get(qn(f"w:{key}"))
            if val:
                pg[f"margin_{key}_mm"] = round(int(val) / 56.7, 1)
        for key in ("header", "footer"):
            val = pgMar.get(qn(f"w:{key}"))
            if val:
                pg[f"{key}_mm"] = round(int(val) / 56.7, 1)
    return pg


def extract_from_notice(notice_path) -> dict:
    """返回 {"page": {...}, "samples": {role: {...}}, "sign": {...} 或 {}}。"""
    doc = Document(str(notice_path))
    result = {"page": _page_from_notice(doc), "samples": {}, "sign": {}}

    paras = list(doc.paragraphs)
    markers = dict(_SAMPLE_MARKERS)
    found = {}
    for idx, p in enumerate(paras):
        text = (p.text or "").strip()
        if not text:
            continue
        for prefix, role in _SAMPLE_MARKERS:
            if role not in found and text.startswith(prefix):
                found[role] = _para_format(p)
                if role == "date":
                    # 日期上一非空段 = 署名样例
                    j = idx - 1
                    while j >= 0 and not (paras[j].text or "").strip():
                        j -= 1
                    if j >= 0:
                        result["sign"] = _para_format(paras[j])
                break
    result["samples"] = found
    return result


def apply_import(rules: dict, notice_path) -> tuple:
    """把通知样例/页面设置合并进规则，返回 (new_rules, report_lines)。"""
    extracted = extract_from_notice(notice_path)
    new_rules = copy.deepcopy(rules)
    report = []

    page = extracted["page"]
    for key, val in page.items():
        if key in new_rules.get("page", {}):
            new_rules["page"][key] = val
    if page:
        report.append(f"[页面] 已从通知导入 {len(page)} 项页面设置")

    for role, fmt in extracted["samples"].items():
        if role not in rules_config.ROLES:
            continue
        label = rules_config.ROLE_LABELS.get(role, role)
        if "font" in fmt:
            new_rules["fonts"][role] = fmt["font"]
        if "size_pt" in fmt:
            new_rules["sizes_pt"][role] = fmt["size_pt"]
        if "align" in fmt:
            new_rules["alignment"][role] = fmt["align"]
        if "line_pt" in fmt:
            new_rules["line_spacing_pt"][role] = fmt["line_pt"]
        desc = "、".join(f"{k}={v}" for k, v in fmt.items())
        report.append(f"[{label}] {desc}")

    if extracted["sign"] and "font" in extracted["sign"]:
        new_rules["fonts"]["sign"] = extracted["sign"]["font"]
    if extracted["sign"] and "size_pt" in extracted["sign"]:
        new_rules["sizes_pt"]["sign"] = extracted["sign"]["size_pt"]

    return new_rules, report


def validate_against_notice(rules: dict, notice_path) -> list:
    """比对通知样例与当前配置，返回差异说明列表。"""
    extracted = extract_from_notice(notice_path)
    report = []
    for role, fmt in extracted["samples"].items():
        if role not in rules_config.ROLES:
            continue
        label = rules_config.ROLE_LABELS.get(role, role)
        diffs = []
        if "font" in fmt and rules["fonts"].get(role) != fmt["font"]:
            diffs.append(f"字体 通知={fmt['font']} 配置={rules['fonts'].get(role)}")
        if "size_pt" in fmt and rules["sizes_pt"].get(role) != fmt["size_pt"]:
            diffs.append(f"字号 通知={fmt['size_pt']} 配置={rules['sizes_pt'].get(role)}")
        if "align" in fmt and rules["alignment"].get(role) != fmt["align"]:
            diffs.append(f"对齐 通知={fmt['align']} 配置={rules['alignment'].get(role)}")
        if "line_pt" in fmt and rules["line_spacing_pt"].get(role) != fmt["line_pt"]:
            diffs.append(f"行距 通知={fmt['line_pt']} 配置={rules['line_spacing_pt'].get(role)}")
        if diffs:
            report.append(f"[{label}] " + "；".join(diffs))
    if not report:
        report.append("当前配置与通知样例完全一致。")
    return report
