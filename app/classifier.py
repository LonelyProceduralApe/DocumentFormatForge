"""结构识别分类器：将段落归类为公文要素（半自动，结果供人工确认）。"""
from __future__ import annotations

import re

# 层次标题识别（顺序重要：先匹配更具体的括号式）
_HEADING_PATTERNS = [
    ("h4", re.compile(r"^（\s*\d+\s*）")),
    ("h3", re.compile(r"^\s*\d+\s*[.．、]")),
    ("h2", re.compile(r"^（\s*[一二三四五六七八九十百]+\s*）")),
    ("h1", re.compile(r"^\s*[一二三四五六七八九十百]+\s*、")),
]

_DATE_RE = re.compile(
    r"^\s*[0-9０-９×X]{2,4}\s*年\s*[0-9０-９×X]{1,2}\s*月\s*[0-9０-９×X]{1,2}\s*日"
)

# 机关单位关键词（用于无冒号主送机关的弱识别）
_ORG_KEYWORDS = (
    "省", "市", "县", "区", "镇", "乡", "部", "委", "办", "局", "厅", "处", "科",
    "公司", "集团", "中心", "学院", "学校", "医院", "银行", "研究院", "所", "队",
    "站", "厂", "集团军", "部委",
)


def classify(paragraphs):
    """paragraphs: [{"idx", "text", "has_pict"}]。返回 {idx: role}。"""
    roles = {}

    # 1) 版记：含"印发"且以"印发"结尾 → 删除
    for p in paragraphs:
        if "印发" in p["text"] and p["text"].rstrip().endswith("印发"):
            roles[p["idx"]] = "delete"

    # 2) 版头发文字号：〔年份〕…号 + 签发人 → 删除
    for p in paragraphs:
        t = p["text"]
        if re.search(r"〔\d{4}〕", t) and "号" in t and "签发人" in t:
            roles[p["idx"]] = "delete"

    # 3) 版头红头/红线：第一个有效文本段之前的 pict 段 → 删除
    first_valid = None
    for p in paragraphs:
        if p["text"] and roles.get(p["idx"]) != "delete":
            first_valid = p["idx"]
            break
    if first_valid is not None:
        for p in paragraphs:
            if p["idx"] < first_valid and p.get("has_pict"):
                roles[p["idx"]] = "delete"

    # 4) 正常分类（有效段落 = 非删除的文本段）
    valid = [p for p in paragraphs if p["text"] and roles.get(p["idx"]) != "delete"]
    if not valid:
        return roles

    title_idx = valid[0]["idx"]
    roles[title_idx] = "title"

    after_title = valid[1:]
    for p in after_title:
        text = p["text"]
        heading = _heading_role(text)
        if heading:
            roles[p["idx"]] = heading
        elif text.startswith("附件："):
            roles[p["idx"]] = "attachment_note"
        elif text.startswith("附件"):
            roles[p["idx"]] = "attachment"
        elif _DATE_RE.match(text):
            roles[p["idx"]] = "date"
        elif text.startswith(("抄送", "主送")):
            roles[p["idx"]] = "copy"
        else:
            roles[p["idx"]] = "body"

    # 附件条目：附件说明之后紧邻的 "N." 行 → 对齐"1."
    for i, p in enumerate(valid):
        if roles.get(p["idx"]) == "attachment_note":
            for q in valid[i + 1:]:
                if re.match(r"^\s*\d+\s*[.．、]", q["text"]):
                    roles[q["idx"]] = "attachment_item"
                else:
                    break

    # 5) 主送机关：标题后第一个被判为 body 的段，若形似主送机关则改判
    first_body = next((p for p in after_title if roles.get(p["idx"]) == "body"), None)
    if first_body is not None and _looks_addressee(first_body["text"]):
        roles[first_body["idx"]] = "addressee"

    # 6) 署名 = 每个成文日期前紧邻的非空段；附注 = 日期后紧邻的括号段
    for p in valid:
        if roles.get(p["idx"]) == "date":
            prev = _prev_nonempty(valid, p)
            if prev is not None and roles.get(prev["idx"]) in ("body", None):
                roles[prev["idx"]] = "sign"
            nxt = _next_nonempty(valid, p)
            if nxt is not None and roles.get(nxt["idx"]) == "body" \
                    and nxt["text"].startswith(("（", "(")):
                roles[nxt["idx"]] = "note"

    return roles


def _heading_role(text: str):
    for role, pat in _HEADING_PATTERNS:
        if pat.match(text):
            return role
    return None


def _looks_addressee(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    if t.endswith(("：", ":")):
        return True
    if "、" in t and len(t) <= 40:
        return True
    if len(t) <= 30 and any(k in t for k in _ORG_KEYWORDS):
        return True
    return False


def _prev_nonempty(nonempty, p):
    idx = nonempty.index(p)
    return nonempty[idx - 1] if idx > 0 else None


def _next_nonempty(nonempty, p):
    idx = nonempty.index(p)
    return nonempty[idx + 1] if idx + 1 < len(nonempty) else None
