import os
import subprocess
import sys
from pathlib import Path


def get_platform_name():
    if sys.platform == "win32":
        return _windows_desktop()

    elif sys.platform == "darwin":
        return _macOS_desktop()

    elif sys.platform.startswith("linux"):
        return _linus_desktop()

    return Path.home() / "Desktop"


def _windows_desktop() -> Path:
    import winreg

    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders",
    ) as key:
        value, _ = winreg.QueryValueEx(key, "Desktop")
    return Path(value)


def _linus_desktop():
    xdg = os.environ.get("XDG_DESKTOP_DIR")
    if xdg:
        return Path(xdg)

    # 2) 读用户配置文件
    cfg = Path.home() / ".config" / "user-dirs.dirs"
    if cfg.exists():
        for line in cfg.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if line.startswith("XDG_DESKTOP_DIR="):
                val = line.split("=", 1)[1].strip().strip('"')
                # 里面可能是 $HOME/Desktop，要展开变量
                return Path(os.path.expandvars(val))

    # 3) 调 xdg-user-dir 命令（多数桌面发行版都有）
    try:
        out = subprocess.check_output(["xdg-user-dir", "DESKTOP"], text=True).strip()
        if out and out != str(Path.home()):
            return Path(out)
    except (FileNotFoundError, subprocess.CalledProcessError):
        pass

    # 4) 兜底
    return Path.home() / "Desktop"


def _macOS_desktop():
    return Path.home() / "Desktop"
