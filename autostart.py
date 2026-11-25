# -*- coding: utf-8 -*-
"""
开机自启动

Windows：HKCU\\...\\Run（不需要管理员权限）
macOS  ：~/Library/LaunchAgents/com.ppttoolbox.autostart.plist（launchd，不需要 sudo）

- 只写当前用户，不动系统级配置
- 每次启动都会核对已写入的路径：用户把程序挪了位置、换了电脑，
  这里会自动改写成当前路径，避免"开机自启动指向不存在的文件"
"""
import os
import sys
import plistlib

import paths

RUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
VALUE_NAME = "PPT工具箱"

MAC_LABEL = "com.ppttoolbox.autostart"


# --------------------------------------------------------------------------- #
# 命令行
# --------------------------------------------------------------------------- #
def _command():
    """返回应该写进系统自启动项的完整命令行（路径带引号，防止空格截断）"""
    exe = paths.exe_path()
    if paths.is_frozen() and exe and exe.lower().endswith(".exe"):
        return '"%s" --autostart' % exe
    if paths.is_frozen():
        # macOS：PyInstaller 的 .app 里真正的可执行文件在 Contents/MacOS 下
        return '"%s" --autostart' % exe
    # 源码运行
    py = sys.executable or ("python3" if paths.MACOS else "python.exe")
    if not paths.MACOS and py.lower().endswith("python.exe"):
        # 源码运行：优先 pythonw.exe，避免弹黑框
        pyw = py[:-len("python.exe")] + "pythonw.exe"
        if os.path.isfile(pyw):
            py = pyw
    script = os.path.join(paths.exe_dir(), "main.py")
    return '"%s" "%s" --autostart' % (py, script)


def _argv():
    """命令行拆成 argv 列表（launchd 需要分开写）"""
    cmd = _command().strip()
    out, cur, quote = [], "", None
    for ch in cmd:
        if quote:
            if ch == quote:
                quote = None
            else:
                cur += ch
        elif ch in ("'", '"'):
            quote = ch
        elif ch == " ":
            if cur:
                out.append(cur)
                cur = ""
        else:
            cur += ch
    if cur:
        out.append(cur)
    return out


# --------------------------------------------------------------------------- #
# macOS：launchd
# --------------------------------------------------------------------------- #
def _mac_plist_path():
    d = os.path.join(os.path.expanduser("~"), "Library", "LaunchAgents")
    paths.ensure_dir(d)
    return os.path.join(d, MAC_LABEL + ".plist")


def _mac_read():
    p = _mac_plist_path()
    if not paths.is_file(p):
        return None
    try:
        with open(p, "rb") as f:
            return plistlib.load(f)
    except Exception:
        return None


def _mac_write(argv):
    data = {
        "Label": MAC_LABEL,
        "ProgramArguments": list(argv),
        "RunAtLoad": True,
        "ProcessType": "Interactive",
        "LimitLoadToSessionType": "Aqua",
    }
    p = _mac_plist_path()
    with open(p, "wb") as f:
        plistlib.dump(data, f)
    return p


def _launchctl(action, plist_path):
    """launchctl load/unload；失败不算错误（新版系统改用 bootstrap 子命令）"""
    import subprocess
    uid = str(os.getuid())
    cmds = {
        "load": ([["launchctl", "load", "-w", plist_path]],
                 [["launchctl", "bootstrap", "gui/" + uid, plist_path]]),
        "unload": ([["launchctl", "unload", "-w", plist_path]],
                   [["launchctl", "bootout", "gui/" + uid, plist_path]]),
    }
    for argv in cmds[action][0] + cmds[action][1]:
        try:
            subprocess.run(argv, stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL, timeout=10)
        except Exception:
            continue


# --------------------------------------------------------------------------- #
# Windows：注册表
# --------------------------------------------------------------------------- #
def _open(create=False):
    import winreg
    access = winreg.KEY_READ
    if create:
        access = winreg.KEY_READ | winreg.KEY_WRITE
    return winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY, 0, access)


# --------------------------------------------------------------------------- #
# 统一接口
# --------------------------------------------------------------------------- #
def current_command():
    """系统里现在写的自启动命令；没有则返回空串。"""
    if paths.MACOS:
        data = _mac_read()
        if not data:
            return ""
        argv = data.get("ProgramArguments") or []
        return " ".join('"%s"' % a if " " in str(a) else str(a) for a in argv)
    try:
        import winreg
        with _open() as key:
            value, _ = winreg.QueryValueEx(key, VALUE_NAME)
            return str(value or "")
    except Exception:
        return ""


def is_enabled():
    return bool(current_command().strip())


def enable():
    """写入自启动项，返回 (成功, 说明)"""
    if paths.MACOS:
        try:
            argv = _argv()
            p = _mac_write(argv)
            _launchctl("load", p)
            return True, " ".join(argv)
        except Exception as exc:
            return False, str(exc)
    try:
        import winreg
        cmd = _command()
        with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, RUN_KEY, 0,
                                winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_SZ, cmd)
        return True, cmd
    except Exception as exc:
        return False, str(exc)


def disable():
    if paths.MACOS:
        try:
            p = _mac_plist_path()
            if paths.is_file(p):
                _launchctl("unload", p)
                os.remove(p)
            return True, ""
        except Exception as exc:
            return False, str(exc)
    try:
        import winreg
        with _open(create=True) as key:
            winreg.DeleteValue(key, VALUE_NAME)
        return True, ""
    except FileNotFoundError:
        return True, ""
    except Exception as exc:
        return False, str(exc)


def sync():
    """自启动开着但路径已经变了（换目录/换电脑）→ 就地修正。返回是否修正过。"""
    if not is_enabled():
        return False
    want = _command()
    have = current_command()
    if _norm(have) == _norm(want):
        return False
    ok, _ = enable()
    return bool(ok)


def _norm(cmd):
    """比较用：去掉多余空白和包裹引号"""
    return " ".join(str(cmd or "").replace('"', " ").split())


def state():
    """给 UI 用的完整状态"""
    have = current_command()
    want = _command()
    return {
        "enabled": bool(have.strip()),
        "command": have,
        "expected": want,
        "stale": bool(have.strip()) and _norm(have) != _norm(want),
        "platform": "macOS" if paths.MACOS else ("Windows" if paths.WINDOWS else os.name),
    }


if __name__ == "__main__":
    import json
    print(json.dumps(state(), ensure_ascii=False, indent=2))
