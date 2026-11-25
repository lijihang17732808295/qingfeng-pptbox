"""Windows 原生文件拖放：直接调 DragQueryFileW（UTF-16）。

不用 windnd 的原因：它走 ANSI 版 DragQueryFile + 260 字节定长缓冲，
一旦进程 ACP 不是 GBK（例如清单里声明了 UTF-8 代码页），中文路径就会
被错误解码成乱码，长路径还会被截断，表现为"文件名乱码 / 找不到文件"。
这里用 W 接口，拿到的是原生 Unicode 字符串，与代码页无关。
"""

import ctypes
from ctypes import wintypes

WM_DROPFILES = 0x0233
GWL_WNDPROC = -4

LRESULT = ctypes.c_ssize_t
WPARAM = ctypes.c_size_t
LPARAM = ctypes.c_ssize_t

_user32 = ctypes.WinDLL("user32", use_last_error=True)
_shell32 = ctypes.WinDLL("shell32", use_last_error=True)

WNDPROC = ctypes.WINFUNCTYPE(LRESULT, wintypes.HWND, wintypes.UINT, WPARAM, LPARAM)

if ctypes.sizeof(ctypes.c_void_p) == 8:
    _GetWindowLong = _user32.GetWindowLongPtrW
    _SetWindowLong = _user32.SetWindowLongPtrW
else:
    _GetWindowLong = _user32.GetWindowLongW
    _SetWindowLong = _user32.SetWindowLongW

_GetWindowLong.restype = ctypes.c_void_p
_GetWindowLong.argtypes = [wintypes.HWND, ctypes.c_int]
_SetWindowLong.restype = ctypes.c_void_p
_SetWindowLong.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_void_p]

_CallWindowProc = _user32.CallWindowProcW
_CallWindowProc.restype = LRESULT
_CallWindowProc.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                            WPARAM, LPARAM]

_DragAcceptFiles = _shell32.DragAcceptFiles
_DragAcceptFiles.restype = None
_DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]

# 注意：GetParent 默认返回 c_int，在 64 位下会把句柄截断成 32 位垃圾值
_GetParent = _user32.GetParent
_GetParent.restype = wintypes.HWND
_GetParent.argtypes = [wintypes.HWND]

_GetAncestor = _user32.GetAncestor
_GetAncestor.restype = wintypes.HWND
_GetAncestor.argtypes = [wintypes.HWND, ctypes.c_uint]
GA_ROOT = 2

_DragQueryFileW = _shell32.DragQueryFileW
_DragQueryFileW.restype = wintypes.UINT
_DragQueryFileW.argtypes = [ctypes.c_void_p, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]

_DragFinish = _shell32.DragFinish
_DragFinish.restype = None
_DragFinish.argtypes = [ctypes.c_void_p]

_hooks = []          # 保活：回调被 GC 回收会导致崩溃
_hooked_hwnds = set()


def log(msg):
    """设 PPTDEBUG=1 时写日志，便于排查打包后的拖放问题"""
    import os
    if not os.environ.get("PPTDEBUG"):
        return
    try:
        import paths
        p = paths.log_file("ppt_drop.log") or os.path.join(
            os.environ.get("TEMP", "."), "ppt_drop.log")
        with paths.open_binary(p, "a", encoding="utf-8") as f:
            f.write("%s\n" % msg)
    except Exception:
        pass


def query_files(hdrop):
    """从 HDROP 取出全部文件的完整路径（Unicode）"""
    count = _DragQueryFileW(hdrop, 0xFFFFFFFF, None, 0)
    out = []
    for i in range(count or 0):
        length = _DragQueryFileW(hdrop, i, None, 0)
        if not length:
            continue
        buf = ctypes.create_unicode_buffer(length + 1)
        _DragQueryFileW(hdrop, i, buf, length + 1)
        if buf.value:
            out.append(buf.value)
    return out


class _Hook:
    def __init__(self, hwnd, callback):
        self.hwnd = int(hwnd)
        self.callback = callback
        self._proc = WNDPROC(self._wndproc)
        _DragAcceptFiles(self.hwnd, True)
        self._old = _GetWindowLong(self.hwnd, GWL_WNDPROC)
        _SetWindowLong(self.hwnd, GWL_WNDPROC,
                       ctypes.cast(self._proc, ctypes.c_void_p).value)

    def _wndproc(self, hwnd, msg, wp, lp):
        if msg == WM_DROPFILES:
            try:
                files = query_files(wp)
                if files:
                    self.callback(files)
            except Exception:
                pass
            try:
                _DragFinish(wp)
            except Exception:
                pass
            return 0
        return _CallWindowProc(self._old, hwnd, msg, wp, lp)


def hook(target, callback):
    """给 tkinter 窗口挂拖放。target 可以是 widget 或 hwnd。"""
    hwnd = target.winfo_id() if hasattr(target, "winfo_id") else int(target)
    # Tk 顶层窗口外面还有一层 wrapper（TkTopLevel / TkChild），整条链都注册
    cands = []

    def add(h):
        if h and int(h) not in cands:
            cands.append(int(h))

    add(hwnd)
    try:
        add(_GetAncestor(wintypes.HWND(hwnd), GA_ROOT))
    except Exception:
        pass
    try:
        top = cands[-1] if cands else hwnd
        p = _GetParent(wintypes.HWND(top))
        add(p)
        if p:
            add(_GetParent(wintypes.HWND(p)))
    except Exception:
        pass
    log("hook candidates: %s" % [hex(int(c)) for c in cands])

    ok = 0
    for h in cands:
        if h in _hooked_hwnds:
            continue
        try:
            _hooks.append(_Hook(h, callback))
            _hooked_hwnds.add(h)
            ok += 1
        except Exception:
            continue
    return ok
