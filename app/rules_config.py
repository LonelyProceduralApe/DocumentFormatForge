"""规则配置加载与保存。

规则以 rules.json 为载体。加载时以「随程序打包的 rules.json」为基线，
再叠加「用户可编辑的 rules.json」（开发环境两者为同一文件；打包成 exe 后
用户版位于 exe 同目录，便于修改）。
"""
from __future__ import annotations

import copy
import json
import sys
from pathlib import Path

# 段落角色（不含页码，页码走页脚）
ROLES = [
    "title", "addressee", "body", "h1", "h2", "h3", "h4",
    "attachment_note", "attachment", "attachment_item", "sign", "date", "note", "copy", "print",
]

ROLE_LABELS = {
    "title": "标题",
    "addressee": "主送机关",
    "body": "正文",
    "h1": "一级标题（一、）",
    "h2": "二级标题（（一））",
    "h3": "三级标题（1.）",
    "h4": "四级标题（（1））",
    "attachment_note": "附件说明",
    "attachment": "附件",
    "attachment_item": "附件条目",
    "sign": "发文机关署名",
    "date": "成文日期",
    "note": "附注",
    "copy": "抄送/主送（版记）",
    "print": "印发机关和日期",
}


def bundled_rules_path() -> Path:
    return Path(__file__).resolve().parent / "rules.json"


def user_rules_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent / "rules.json"
    return bundled_rules_path()


def _deep_merge(base: dict, extra: dict) -> None:
    for key, value in extra.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            _deep_merge(base[key], value)
        else:
            base[key] = value


# 最小可用默认（当 rules.json 缺失/损坏时的兜底，避免程序崩溃）
_MINIMAL = {
    "page": {"width_mm": 210, "height_mm": 297, "margin_top_mm": 37, "margin_bottom_mm": 34,
             "margin_left_mm": 28, "margin_right_mm": 26, "gutter_mm": 0, "header_mm": 15,
             "footer_mm": 30, "lines_per_page": 22, "chars_per_line": 28},
    "fonts": {"title": "方正小标宋简体", "addressee": "仿宋_GB2312", "body": "仿宋_GB2312",
              "h1": "黑体", "h2": "楷体_GB2312", "h3": "仿宋_GB2312", "h4": "仿宋_GB2312",
              "attachment_note": "仿宋_GB2312", "attachment": "黑体", "attachment_item": "仿宋_GB2312",
               "sign": "仿宋_GB2312", "date": "仿宋_GB2312",
              "note": "仿宋_GB2312", "copy": "仿宋_GB2312", "print": "仿宋_GB2312",
              "page_number": "宋体", "ascii": "Times New Roman"},
    "sizes_pt": {"title": 22, "addressee": 16, "body": 16, "h1": 16, "h2": 16, "h3": 16, "h4": 16,
                 "attachment_note": 16, "attachment": 16, "attachment_item": 16, "sign": 16, "date": 16, "note": 16,
                 "copy": 14, "print": 14, "page_number": 14},
    "line_spacing_pt": {"title": 35, "addressee": 29, "body": 29, "h1": 29, "h2": 29, "h3": 29, "h4": 29,
                        "attachment_note": 29, "attachment": 29, "attachment_item": 29, "sign": 29, "date": 29, "note": 29,
                        "copy": 29, "print": 29},
    "alignment": {"title": "center", "addressee": "left", "body": "justify", "h1": "justify",
                  "h2": "justify", "h3": "justify", "h4": "justify", "attachment_note": "left",
                  "attachment": "left", "attachment_item": "left", "sign": "right", "date": "right", "note": "left",
                  "copy": "left", "print": "left"},
    "indent": {"body_first_line_chars": 2, "heading_first_line_chars": 2, "attachment_first_line_chars": 2, "attachment_item_left_chars": 5,
               "note_first_line_chars": 2, "date_right_chars": 5, "sign_right_chars": 6,
               "copy_left_right_chars": 1, "print_left_right_chars": 1},
    "page_number": {"enabled": True, "align": "justify", "left_chars": 1, "right_chars": 1,
                    "line_spacing": "single", "line_spacing_value": 1.0,
                    "dash_before": "\u2014", "dash_after": "\u2014", "space": " "},
    "output": {"suffix": "_已排版"},
}


def load_rules(path=None) -> dict:
    """加载规则：以内置最小默认兜底，再叠加打包版、用户版。"""
    rules = copy.deepcopy(_MINIMAL)
    bundled = bundled_rules_path()
    if bundled.exists():
        try:
            with open(bundled, "r", encoding="utf-8") as f:
                _deep_merge(rules, json.load(f))
        except Exception:
            pass
    user = Path(path) if path else user_rules_path()
    if user.exists() and user.resolve() != bundled.resolve():
        try:
            with open(user, "r", encoding="utf-8") as f:
                _deep_merge(rules, json.load(f))
        except Exception:
            pass
    return rules


def save_rules(rules: dict, path=None) -> Path:
    target = Path(path) if path else user_rules_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "w", encoding="utf-8") as f:
        json.dump(rules, f, ensure_ascii=False, indent=2)
    return target


def role_font(rules: dict, role: str) -> str:
    return rules["fonts"].get(role, rules["fonts"].get("body", "仿宋_GB2312"))


def ascii_font(rules: dict) -> str:
    return rules["fonts"].get("ascii", "Times New Roman")


def role_size_pt(rules: dict, role: str) -> float:
    return float(rules["sizes_pt"].get(role, 16))


def role_line_spacing_pt(rules: dict, role: str) -> float:
    return float(rules["line_spacing_pt"].get(role, 29))


def role_alignment(rules: dict, role: str) -> str:
    return rules["alignment"].get(role, "justify")
