"""测试「从《通知》导入/校验规则」功能。"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from app import notice_import, rules_config

NOTICE = ROOT / "关于进一步规范公文格式的通知.docx"


def main():
    rules = rules_config.load_rules()
    print("=== 从通知提取 ===")
    extracted = notice_import.extract_from_notice(NOTICE)
    print("page:", extracted["page"])
    for role, fmt in extracted["samples"].items():
        print(f"  {role}: {fmt}")
    print("sign:", extracted["sign"])

    print("\n=== 与当前配置比对 ===")
    for line in notice_import.validate_against_notice(rules, NOTICE):
        print("  " + line)


if __name__ == "__main__":
    main()
