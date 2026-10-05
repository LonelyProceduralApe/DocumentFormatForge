"""字体可用性检测。

读取系统级(HKLM)与用户级(HKCU)字体注册表，以及用户级字体目录，
比对规则中要求的字体名是否可用。对「黑体/宋体/仿宋/楷体」等中英文
异名字体做了别名映射。
"""
from __future__ import annotations

import os
import winreg

_FONT_DIR = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")

# 常见中英文异名字体别名（小写）
_ALIASES = {
    "宋体": {"simsun", "nsimsun"},
    "黑体": {"simhei"},
    "仿宋": {"fangsong"},
    "楷体": {"kaiti"},
    "微软雅黑": {"microsoft yahei", "msyh"},
    "仿宋_gb2312": {"fangsong_gb2312"},
    "楷体_gb2312": {"kaiti_gb2312"},
    "times new roman": {"times"},
}


def _strip_suffix(name: str) -> str:
    for suffix in (" (TrueType)", " (OpenType)", " (All res)", " (TrueType collection)", " (OpenType collection)"):
        if name.lower().endswith(suffix.lower()):
            return name[: -len(suffix)]
    return name


def installed_font_names() -> set:
    names = set()
    for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            key = winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts")
            i = 0
            while True:
                try:
                    value_name, _value, _type = winreg.EnumValue(key, i)
                    raw = _strip_suffix(value_name).strip().lower()
                    # 拆分合并条目，如 "SimSun & NSimSun" → simsun / nsimsun
                    for part in raw.split("&"):
                        part = part.strip()
                        if part:
                            names.add(part)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except OSError:
            pass
    # 用户级字体目录（有些字体只放文件、不在注册表）
    if os.path.isdir(_FONT_DIR):
        try:
            for fn in os.listdir(_FONT_DIR):
                names.add(os.path.splitext(fn)[0].strip().lower())
        except OSError:
            pass
    return names


def check_fonts(required_names) -> list:
    """返回缺失字体名列表。"""
    installed = installed_font_names()
    missing = []
    for name in required_names:
        low = name.strip().lower()
        candidates = {low} | _ALIASES.get(low, set())
        hit = False
        for cand in candidates:
            if cand in installed or any(cand in inst or inst in cand for inst in installed):
                hit = True
                break
        if not hit:
            missing.append(name)
    return missing
