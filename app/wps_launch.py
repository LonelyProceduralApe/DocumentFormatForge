"""定位并启动 WPS 打开结果文件。"""
from __future__ import annotations

import glob
import os
import subprocess
import winreg

_COMMON_ROOTS = [
    r"D:\OfficeTool\WPS\WPS Office",
    r"C:\Program Files\Kingsoft\WPS Office",
    r"C:\Program Files (x86)\Kingsoft\WPS Office",
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Kingsoft", "WPS Office"),
    os.path.join(os.environ.get("ProgramFiles", r"C:\Program Files"), "WPS Office"),
]


def _uninstall_entries():
    keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    ]
    for root, path in keys:
        try:
            key = winreg.OpenKey(root, path)
            i = 0
            while True:
                try:
                    sub = winreg.OpenKey(key, winreg.EnumKey(key, i))
                    try:
                        display = winreg.QueryValueEx(sub, "DisplayName")[0]
                    except OSError:
                        display = ""
                    if display and "wps" in display.lower():
                        values = {}
                        for name in ("InstallLocation", "DisplayIcon", "UninstallString"):
                            try:
                                values[name] = winreg.QueryValueEx(sub, name)[0]
                            except OSError:
                                values[name] = ""
                        yield values
                    winreg.CloseKey(sub)
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except OSError:
            continue


def find_wps_exe() -> str | None:
    candidates = set()
    for entry in _uninstall_entries():
        for val in (entry.get("InstallLocation", ""), entry.get("DisplayIcon", ""), entry.get("UninstallString", "")):
            if val and os.path.isabs(val):
                d = os.path.dirname(val)
                # 逐级向上，找含 office6\wps.exe 的目录
                for _ in range(6):
                    candidates.add(os.path.join(d, "office6", "wps.exe"))
                    if os.path.basename(d).lower() in ("wps office", "office6"):
                        break
                    d = os.path.dirname(d)
    for root in _COMMON_ROOTS:
        if root:
            candidates.update(glob.glob(os.path.join(root, "*", "office6", "wps.exe")))
            candidates.update(glob.glob(os.path.join(root, "office6", "wps.exe")))
    for c in sorted(candidates):
        if os.path.isfile(c):
            return c
    return None


def open_with_wps(path) -> bool:
    """用 WPS 打开文件；失败则退回系统默认关联。"""
    path = os.path.abspath(str(path))
    exe = find_wps_exe()
    if exe:
        try:
            subprocess.Popen(
                [exe, path],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                stdin=subprocess.DEVNULL,
                close_fds=True,
            )
            return True
        except Exception:
            pass
    try:
        os.startfile(path)  # noqa: PTH? 兜底
        return True
    except Exception:
        return False
