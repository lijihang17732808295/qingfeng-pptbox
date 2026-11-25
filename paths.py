# -*- coding: utf-8 -*-
"""
统一路径层（换电脑 / 换用户 / 长路径 / 中文路径 / 盘符变化 的兜底）

设计原则
--------
1. 不假设盘符、用户名、Program Files 位置、当前工作目录；
2. 资源路径按候选列表逐个探测：_MEIPASS -> exe 目录 -> exe/_internal
   -> 源码目录 -> 当前目录（打包 onefile / onedir / 源码运行都成立）；
3. 所有"可选路径"先校验再使用，失败有回退，绝不让一个路径把整个功能带崩；
4. 用户可写目录统一走 Known Folder（SHGetKnownFolderPath），拿不到再退环境变量；
5. 超过 260 的路径在做文件 IO 时自动加 \\\\?\\ 前缀，交给 PowerPoint COM 时再剥掉；
6. 用户拖进来 / 粘贴进来的路径统一走 resolve_existing()，能修就修，修不了给出可执行的提示。
"""
import os
import re
import sys
import time
import ctypes
import tempfile
import threading
from urllib.parse import unquote, urlparse

WINDOWS = os.name == "nt"
MACOS = sys.platform == "darwin"

# --------------------------------------------------------------------------- #
# Known Folder GUID（只在 Windows 上用；拿不到就走环境变量）
# --------------------------------------------------------------------------- #
FOLDERID_LOCALAPPDATA = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
FOLDERID_ROAMINGAPPDATA = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
FOLDERID_FONTS = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
FOLDERID_WINDOWS = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
FOLDERID_PROFILE = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
FOLDERID_DOCUMENTS = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"

_GUID_STRUCTS = {}


class _GUID(ctypes.Structure):
    _fields_ = [("Data1", ctypes.c_ulong),
                ("Data2", ctypes.c_ushort),
                ("Data3", ctypes.c_ushort),
                ("Data4", ctypes.c_ubyte * 8)]


def _guid(text):
    """把 "{XXXXXXXX-XXXX-...}" 转成 GUID 结构（缓存，避免重复解析）"""
    g = _GUID_STRUCTS.get(text)
    if g is not None:
        return g
    s = text.strip().strip("{}")
    a, b, c, d, e = s.split("-")
    g = _GUID()
    g.Data1 = int(a, 16)
    g.Data2 = int(b, 16)
    g.Data3 = int(c, 16)
    for i, v in enumerate(bytes.fromhex(d + e)):
        g.Data4[i] = v
    _GUID_STRUCTS[text] = g
    return g


_MAC_FOLDERS = {
    FOLDERID_DOCUMENTS: ("Documents",),
    FOLDERID_PROFILE: (),
    FOLDERID_ROAMINGAPPDATA: ("Library", "Application Support"),
    FOLDERID_LOCALAPPDATA: ("Library", "Application Support"),
}


def known_folder(guid, fallback=None):
    """取系统 Known Folder；失败返回 fallback。

    注意：SHGetKnownFolderPath 在 shell32 里，不在 ole32（早前写成 ole32
    会一直静默失败，只能退回环境变量，遇到 OneDrive 重定向就取错目录）。
    macOS 没有 Known Folder，按 GUID 映射到 ~ 下的等价目录。
    """
    if MACOS:
        parts = _MAC_FOLDERS.get(guid)
        if parts is None:
            return fallback
        return os.path.join(os.path.expanduser("~"), *parts) if parts \
            else os.path.expanduser("~")
    if not WINDOWS:
        return fallback
    g = _guid(guid)
    for lib in ("shell32", "ole32"):
        try:
            fn = getattr(ctypes.windll, lib).SHGetKnownFolderPath
        except Exception:
            continue
        try:
            fn.restype = ctypes.c_long
            fn.argtypes = [ctypes.POINTER(_GUID), ctypes.c_ulong,
                           ctypes.c_void_p, ctypes.POINTER(ctypes.c_wchar_p)]
        except Exception:
            pass
        ptr = ctypes.c_wchar_p()
        try:
            hr = fn(ctypes.byref(g), 0, None, ctypes.byref(ptr))
            if hr == 0 and ptr.value:
                value = ptr.value
                try:
                    ctypes.windll.ole32.CoTaskMemFree(
                        ctypes.cast(ptr, ctypes.c_void_p))
                except Exception:
                    pass
                return value
        except Exception:
            pass
    return fallback


# --------------------------------------------------------------------------- #
# 基础清洗
# --------------------------------------------------------------------------- #
_QUOTE_PAIRS = (('"', '"'), ("'", "'"), ("\u201c", "\u201d"), ("\u2018", "\u2019"))


def clean(p):
    """去掉引号/空白/BOM，把 bytes 解成 str，把 file:// 还原成路径。"""
    if p is None:
        return ""
    if isinstance(p, bytes):
        for enc in ("utf-8", "gbk", "mbcs", "latin-1"):
            try:
                p = p.decode(enc)
                break
            except Exception:
                continue
        else:
            p = p.decode("utf-8", "surrogateescape")
    p = str(p).replace("\ufeff", "").strip()
    # 资源管理器"复制为路径"带双引号；从聊天软件粘贴常带全角引号
    for left, right in _QUOTE_PAIRS:
        if len(p) >= 2 and p.startswith(left) and p.endswith(right):
            p = p[1:-1].strip()
            break
    if p.lower().startswith("file:"):
        try:
            u = urlparse(p)
            rest = unquote(u.path)
            if u.netloc:
                rest = "//" + u.netloc + rest
            if WINDOWS and re.match(r"^/[A-Za-z]:", rest):
                rest = rest[1:]          # file:///C:/x -> C:/x
            p = rest
        except Exception:
            pass
    return p


def expand(p):
    """展开 %VAR%、~；未知变量原样保留。"""
    if not p:
        return p
    try:
        p = os.path.expandvars(p)
    except Exception:
        pass
    if p.startswith("~"):
        try:
            p = os.path.expanduser(p)
        except Exception:
            pass
    return p


def normalize(p, base=None):
    """得到干净的原生路径（保留相对路径语义，除非给了 base）。"""
    p = expand(clean(p))
    if not p:
        return ""
    if p.startswith("\\\\?\\"):
        return p
    if WINDOWS:
        p = p.replace("/", "\\")
    if base and not os.path.isabs(p):
        try:
            p = os.path.join(base, p)
        except Exception:
            pass
    try:
        p = os.path.normpath(p)
    except Exception:
        pass
    return p


def abspath(p, base=None):
    """绝不因 cwd 失效而抛异常的 abspath。"""
    p = normalize(p, base)
    if not p:
        return p
    try:
        return os.path.abspath(p)
    except Exception:
        if re.match(r"^[A-Za-z]:\\", p) or p.startswith("\\\\"):
            return p
        return p


# --------------------------------------------------------------------------- #
# 长路径
# --------------------------------------------------------------------------- #
def long_path(p):
    """文件 IO 用的 \\\\?\\ 形式；非 Windows 或相对路径原样返回。"""
    if not WINDOWS:
        return p
    p = normalize(p)
    if not p or p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    if re.match(r"^[A-Za-z]:\\", p):
        return "\\\\?\\" + p
    return p


def short_path(p):
    """8.3 短路径（给不认 \\\\?\\ / 超长路径的第三方组件用）；失败原样返回。"""
    if not WINDOWS:
        return p
    p = normalize(p)
    try:
        from ctypes import wintypes
        get = ctypes.windll.kernel32.GetShortPathNameW
        get.restype = wintypes.DWORD
        get.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        size = 512
        for _ in range(4):
            buf = ctypes.create_unicode_buffer(size)
            n = get(p, buf, size)
            if 0 < n < size:
                return buf.value
            if n >= size:
                size = n + 1
                continue
            break
    except Exception:
        pass
    return p


def long_name_path(p):
    """8.3 -> 长文件名；失败原样返回。"""
    if not WINDOWS:
        return p
    p = normalize(p)
    try:
        from ctypes import wintypes
        get = ctypes.windll.kernel32.GetLongPathNameW
        get.restype = wintypes.DWORD
        get.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
        size = 512
        for _ in range(4):
            buf = ctypes.create_unicode_buffer(size)
            n = get(p, buf, size)
            if 0 < n < size:
                return buf.value
            if n >= size:
                size = n + 1
                continue
            break
    except Exception:
        pass
    return p


# --------------------------------------------------------------------------- #
# 探测
# --------------------------------------------------------------------------- #
def _probe(fn, p):
    try:
        if fn(p):
            return True
    except Exception:
        pass
    if WINDOWS:
        try:
            if fn(long_path(p)):
                return True
        except Exception:
            pass
    return False


def exists(p):
    return _probe(os.path.exists, p)


def is_file(p):
    return _probe(os.path.isfile, p)


def is_dir(p):
    return _probe(os.path.isdir, p)


def _decode_variants(p):
    """把被错误解码的路径还原（UTF-8 当 GBK 解、GBK 当 UTF-8 解等）。"""
    out = []
    for enc in ("gbk", "utf-8", "big5", "mbcs"):
        try:
            out.append(p.encode(enc).decode("utf-8"))
        except Exception:
            pass
        try:
            out.append(p.encode("latin-1").decode(enc))
        except Exception:
            pass
    return out


def _ci_path(p):
    """按段做大小写不敏感查找（顺便吃掉段尾多出来的空格/点）。"""
    if not WINDOWS:
        return p
    p = normalize(p)
    drive, rest = os.path.splitdrive(p)
    if not rest:
        return p
    if drive and not rest.startswith("\\"):
        return p                       # "C:foo" 这类驱动器相对路径不猜
    cur = drive + "\\" if drive else ("" if not rest.startswith("\\") else "\\")
    for part in [x for x in rest.split("\\") if x not in ("", ".")]:
        cand = os.path.join(cur, part)
        if exists(cand):
            cur = cand
            continue
        try:
            entries = os.listdir(long_path(cur))
        except Exception:
            return p
        low = part.rstrip(" .").lower()
        hit = None
        for e in entries:
            if e.lower() == low or e.rstrip(" .").lower() == low:
                hit = e
                break
        if hit is None:
            return p
        cur = os.path.join(cur, hit)
    return cur


def resolve_existing(p, base=None):
    """把用户给的路径尽量还原成真实存在的路径。

    顺序：清洗 -> 直接判定 -> 长短名互换 -> 乱码还原 -> 逐段大小写不敏感。
    修不好就原样返回，让上层做提示（不做静默失败）。
    """
    p = normalize(p, base)
    if not p:
        return p
    if exists(p):
        return p
    for alt in (long_name_path(p), short_path(p)):
        if alt and alt != p and exists(alt):
            return alt
    for alt in _decode_variants(p):
        if alt and alt != p and exists(alt):
            return alt
    alt = _ci_path(p)
    if alt and alt != p and exists(alt):
        return alt
    return p


def same_file(a, b):
    """两个路径是否指向同一个文件（兼容 8.3 / 大小写 / 符号链接）。"""
    a, b = abspath(a), abspath(b)
    if not a or not b:
        return False
    if os.path.normcase(a) == os.path.normcase(b):
        return True
    try:
        return os.path.samefile(long_path(a), long_path(b))
    except Exception:
        return False


def describe_missing(p):
    """给"找不到文件"弹窗用的具体原因，避免只丢一句冷冰冰的路径。"""
    p = normalize(p)
    if not p:
        return ["路径为空。"]
    if exists(p):
        return ["文件存在，但当前无法读取（可能正被其他程序占用，或权限不足）"]
    out = []
    drive, _ = os.path.splitdrive(p)
    if drive and not is_dir(drive + "\\"):
        out.append("盘符 %s 在本机不存在（U 盘/移动硬盘未插、网络盘未映射）" % drive)
    parent = os.path.dirname(p)
    if parent and not is_dir(parent):
        out.append("所在文件夹不存在：%s" % parent)
    if WINDOWS and len(p) >= 240:
        out.append("路径长度 %d 字符，可能被系统的 260 限制截断" % len(p))
    if p.startswith("\\\\"):
        out.append("这是网络共享路径，请确认网络已连接、共享仍可访问")
    low = p.lower()
    if "onedrive" in low or "dropbox" in low or "坚果云" in p:
        out.append("网盘/OneDrive 文件可能只在云端占位，尚未下载到本机")
    if not out:
        out.append("文件可能已被移动、重命名或删除")
        out.append("路径中可能含有无法识别的字符（可尝试改用「选择文件」）")
    return out


def safe_initialdir(p):
    """给 tkinter 文件对话框用的 initialdir：一定存在，否则返回 None。

    （initialdir 指向不存在的目录时，部分 Tk 版本会直接抛
      "couldn't read directory ... no such file or directory"）
    """
    p = normalize(p)
    if not p:
        return None
    if is_dir(p):
        return p
    if is_file(p):
        d = os.path.dirname(p)
        if d and is_dir(d):
            return d
    cur = p
    for _ in range(24):
        parent = os.path.dirname(cur.rstrip("\\/"))
        if not parent or parent == cur:
            break
        cur = parent
        if is_dir(cur):
            return cur
    return None


def ensure_dir(p):
    """建目录；成功（或本来就存在）返回 True。"""
    p = normalize(p)
    if not p:
        return False
    if is_dir(p):
        return True
    for cand in (long_path(p), p):
        try:
            os.makedirs(cand, exist_ok=True)
            if os.path.isdir(cand):
                return True
        except Exception:
            continue
    return False


def is_writable(p):
    """目录是否可写（真正落一个临时文件）。

    判定只看"能否创建"：某些安全软件会拦住随后的删除，
    若把删除失败也算作不可写，会把整个应用误判到便携目录。
    """
    if not is_dir(p):
        return False
    probe = os.path.join(p, ".wbtest_%d_%d.tmp" % (os.getpid(), int(time.time() * 1000)))
    try:
        with open(long_path(probe), "wb") as f:
            f.write(b"0")
    except Exception:
        return False
    try:
        os.remove(long_path(probe))
    except Exception:
        pass
    return True


# --------------------------------------------------------------------------- #
# 应用自身位置与资源
# --------------------------------------------------------------------------- #
def is_frozen():
    return bool(getattr(sys, "frozen", False))


def exe_path():
    """打包后是 exe 本身；源码运行时是入口脚本（不依赖 argv[0]）。"""
    if is_frozen():
        return abspath(sys.executable)
    try:
        here = os.path.abspath(__file__)
        if is_file(here):
            return here
    except Exception:
        pass
    try:
        return abspath(sys.argv[0])
    except Exception:
        return ""


def exe_dir():
    p = exe_path()
    d = os.path.dirname(p) if p else ""
    if d and is_dir(d):
        return d
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return ""


def base_dirs():
    """资源可能所在的所有目录（按优先级，已去重）。"""
    out = []

    def add(d):
        if not d:
            return
        d = abspath(d)
        if not d or not is_dir(d):
            return
        key = os.path.normcase(d)
        for got in out:
            if os.path.normcase(got) == key:
                return
        out.append(d)

    add(getattr(sys, "_MEIPASS", None))       # onefile 解压目录
    add(exe_dir())                            # exe 同目录（onedir / 便携版）
    add(os.path.join(exe_dir(), "_internal"))  # onedir 新版布局
    try:
        add(os.path.dirname(os.path.abspath(__file__)))  # 源码运行
    except Exception:
        pass
    try:
        cwd = os.getcwd()
        if is_dir(cwd):
            add(cwd)
    except Exception:
        pass
    return out


def resource_path(*parts):
    """按候选目录探测资源，返回第一个存在的；都不存在返回首选拼接结果。"""
    rel = os.path.join(*[str(x) for x in parts]) if parts else ""
    bases = base_dirs()
    for base in bases:
        cand = os.path.join(base, rel) if rel else base
        if exists(cand):
            return cand
    if bases:
        return os.path.join(bases[0], rel) if rel else bases[0]
    return rel


def ensure_cwd(preferred=None):
    """当前工作目录失效（旧快捷方式"起始位置"指向已删除的盘）时切到可用目录。"""
    try:
        cwd = os.getcwd()
        if is_dir(cwd):
            return cwd
    except Exception:
        pass
    home = known_folder(FOLDERID_PROFILE, os.path.expanduser("~"))
    for cand in (preferred, exe_dir(), user_data_dir(), home):
        if cand and is_dir(cand):
            try:
                os.chdir(cand)
                return cand
            except Exception:
                continue
    return None


# --------------------------------------------------------------------------- #
# 用户可写目录 / 临时目录
# --------------------------------------------------------------------------- #
_LOCK = threading.RLock()
_MEMO = {}


def _local_appdata():
    if MACOS:
        return os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    return known_folder(FOLDERID_LOCALAPPDATA, os.environ.get("LOCALAPPDATA"))


def _roaming_appdata():
    if MACOS:
        return os.path.join(os.path.expanduser("~"), "Library", "Application Support")
    return known_folder(FOLDERID_ROAMINGAPPDATA, os.environ.get("APPDATA"))


def user_data_dir():
    """应用数据目录：一定返回一个路径，且尽量保证可写。"""
    with _LOCK:
        cached = _MEMO.get("user_data")
        if cached and is_dir(cached):
            return cached
        cands = []

        def add(p):
            if p:
                p = abspath(p)
                if p and p not in cands:
                    cands.append(p)

        la = _local_appdata()
        ra = _roaming_appdata()
        if la:
            add(os.path.join(la, "PPTToolbox"))
        if ra:
            add(os.path.join(ra, "PPTToolbox"))
        try:
            add(os.path.join(tempfile.gettempdir(), "PPTToolbox"))
        except Exception:
            pass
        home = known_folder(FOLDERID_PROFILE, os.path.expanduser("~"))
        if home:
            add(os.path.join(home, "PPTToolbox"))
        add(os.path.join(exe_dir(), "PPTToolbox_data"))

        for c in cands:
            if ensure_dir(c) and is_writable(c):
                _MEMO["user_data"] = c
                return c
        if cands:
            return cands[0]
        return ""


def cache_dir():
    return user_data_dir()


def cache_file(name):
    """缓存文件路径；目录不可写时返回 None（调用方跳过缓存即可）。"""
    d = cache_dir()
    if not d:
        return None
    if not ensure_dir(d):
        return None
    return os.path.join(d, name)


def log_file(name, rotate=512 * 1024):
    """日志文件路径；超过 rotate 字节就先轮转一次。"""
    d = user_data_dir()
    if not d:
        return ""
    p = os.path.join(d, name)
    try:
        if os.path.isfile(p) and os.path.getsize(p) > rotate:
            old = p + ".1"
            try:
                os.replace(p, old)
            except Exception:
                pass
    except Exception:
        pass
    return p


def temp_dir(prefix="pptfx_"):
    """建一个确定可写的临时目录；系统 TEMP 失效时逐级回退。"""
    cset = []

    def add(p):
        if p:
            cset.append(p)

    try:
        add(tempfile.gettempdir())
    except Exception:
        pass
    la = _local_appdata()
    if la:
        add(os.path.join(la, "Temp"))
    ud = user_data_dir()
    if ud:
        add(os.path.join(ud, "tmp"))
    add(exe_dir())

    # 优先用最短的临时目录：COM/Office 对路径长度敏感
    seen = set()
    bases = []
    for p in cset:
        if p and p not in seen:
            seen.add(p)
            bases.append(p)
    bases.sort(key=len)

    errs = []
    for base in bases:
        if not base or not ensure_dir(base) or not is_writable(base):
            continue
        try:
            return tempfile.mkdtemp(prefix=prefix, dir=base)
        except Exception as exc:
            errs.append("%s: %s" % (base, exc))
    raise RuntimeError(
        "找不到可写的临时目录，无法继续。\n\n"
        "已尝试：\n%s\n\n"
        "通常是 TEMP/TMP 环境变量指向了不存在的目录，或安全软件拦截。"
        % ("\n".join(errs) or "（无）"))


def open_binary(p, mode="rb", encoding=None, errors=None):
    """长路径安全的 open（返回文件对象，调用方用 with）。"""
    kw = {}
    if encoding and "b" not in mode:
        kw["encoding"] = encoding
    if errors:
        kw["errors"] = errors
    try:
        return open(long_path(p), mode, **kw)
    except OSError:
        return open(p, mode, **kw)


def read_text(p, encoding="utf-8"):
    with open_binary(p, "r", encoding=encoding, errors="replace") as f:
        return f.read()


def write_text(p, text, encoding="utf-8"):
    d = os.path.dirname(p)
    if d:
        ensure_dir(d)
    with open(long_path(p), "w", encoding=encoding) as f:
        f.write(text)
    return p


# --------------------------------------------------------------------------- #
# 字体目录
# --------------------------------------------------------------------------- #
def font_dirs():
    """本机可能存放字体的目录（只返回真实存在的）。"""
    cands = []

    def add(p):
        if p and p not in cands:
            cands.append(p)

    if MACOS:
        home = os.path.expanduser("~")
        for d in ("/System/Library/Fonts",
                  "/System/Library/Fonts/Supplemental",
                  "/Library/Fonts",
                  os.path.join(home, "Library", "Fonts")):
            add(d)
        return [d for d in cands if is_dir(d)]

    add(known_folder(FOLDERID_FONTS, None))
    add(os.path.join(known_folder(FOLDERID_WINDOWS,
                                  os.environ.get("WINDIR", "C:\\Windows")), "Fonts"))
    la = _local_appdata()
    if la:
        add(os.path.join(la, "Microsoft", "Windows", "Fonts"))
    # 少数机器把字体装在共享目录（注册表里常见）
    for env, sub in (("SystemRoot", "Fonts"), ("WINDIR", "Fonts")):
        base = os.environ.get(env)
        if base:
            add(os.path.join(base, sub))
    return [d for d in cands if is_dir(d)]


def user_fonts_dir():
    """当前用户字体目录（无需管理员即可写入）。"""
    if MACOS:
        d = os.path.join(os.path.expanduser("~"), "Library", "Fonts")
        ensure_dir(d)
        return d
    la = _local_appdata()
    if la:
        d = os.path.join(la, "Microsoft", "Windows", "Fonts")
        if ensure_dir(d):
            return d
    # 退而求其次：系统字体目录（可能需要管理员）
    win = known_folder(FOLDERID_WINDOWS, os.environ.get("WINDIR", "C:\\Windows"))
    return os.path.join(win, "Fonts")


# --------------------------------------------------------------------------- #
# 自检 / 诊断
# --------------------------------------------------------------------------- #
def open_path(p):
    """用系统默认程序打开文件或文件夹（跨平台）。"""
    p = normalize(p)
    if not p or not exists(p):
        return False
    try:
        if WINDOWS:
            os.startfile(p)
            return True
        import subprocess
        opener = "open" if MACOS else "xdg-open"
        subprocess.Popen([opener, p], close_fds=True,
                         stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return True
    except Exception:
        return False


def diagnose():
    """返回 [(项目, 是否OK, 详情)]，用于换机器后的快速排查。"""
    rows = []

    def add(name, ok, detail=""):
        rows.append((name, bool(ok), str(detail)))

    if is_frozen():
        add("程序位置", is_dir(exe_dir()), "%s（%s）" % (exe_dir(), exe_path()))
    else:
        add("程序位置", is_dir(exe_dir()), "%s（源码运行）" % exe_dir())
    try:
        cwd = os.getcwd()
        add("工作目录", is_dir(cwd), cwd)
    except Exception as exc:
        add("工作目录", False, "已失效：%s" % exc)
    add("运行模式", True,
        ("打包 APP" if MACOS else "打包 EXE") if is_frozen() else "源码")
    bases = base_dirs()
    add("资源候选目录", bool(bases), " ; ".join(bases) or "无")
    for rel in (("assets", "logo.png"), ("assets", "logo_mark.png")):
        p = resource_path(*rel)
        add("资源 %s" % rel[1], is_file(p), p)
    ud = user_data_dir()
    add("用户数据目录", is_writable(ud), ud)
    try:
        t = temp_dir("wbcheck_")
        add("临时目录", True, t)
        try:
            os.rmdir(t)
        except Exception:
            pass
    except Exception as exc:
        add("临时目录", False, str(exc))
    dirs = font_dirs()
    add("字体目录", bool(dirs), " | ".join(dirs) or "未找到")
    try:
        import font_detect
        add("字体索引缓存", True, str(font_detect._cache_path()))
    except Exception as exc:
        add("字体索引缓存", False, exc)
    try:
        import slide_export
        if MACOS:
            add("PowerPoint", bool(slide_export.find_powerpoint_exe()),
                slide_export.find_powerpoint_exe() or "未找到 Microsoft PowerPoint")
            add("LibreOffice", bool(slide_export.find_libreoffice()),
                slide_export.find_libreoffice() or "未找到（图片化需要 PowerPoint 或 LibreOffice）")
            add("PDF 渲染组件", slide_export.has_pdf_renderer(),
                "PyMuPDF 可用" if slide_export.has_pdf_renderer() else "PyMuPDF 缺失")
        else:
            exe = slide_export.find_powerpoint_exe()
            add("PowerPoint", bool(exe), exe or "未找到（将回退到 COM 方案）")
    except Exception as exc:
        add("Office", False, exc)
    return rows


def diagnostic_report():
    try:
        import platform
        osname = "%s %s" % (platform.system(), platform.release())
    except Exception:
        osname = os.name
    lines = [
        "PPT 工具箱 · 路径环境自检",
        "时间：%s" % time.strftime("%Y-%m-%d %H:%M:%S"),
        "系统：%s" % osname,
        "Python：%s" % sys.version.split()[0],
        "",
    ]
    for name, ok, detail in diagnose():
        lines.append("[%s] %-12s %s" % ("OK " if ok else "!! ", name, detail))
    return "\n".join(lines) + "\n"


def write_diagnostic_report(path=None):
    text = diagnostic_report()
    target = path or log_file("路径自检.txt", rotate=0)
    if target:
        try:
            write_text(target, text)
        except Exception:
            return "", text
    return target, text
