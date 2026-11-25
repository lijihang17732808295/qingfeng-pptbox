# -*- coding: utf-8 -*-
"""
PPT 页面图片化固定模块（v2）

背景（v2 重写原因）
-------------------
大量装了 WPS 的机器上，注册表里的 PowerPoint.Application 被 WPS 接管：
    HKCR\\CLSID\\{91493441-5A91-11CF-8700-00AA0060263B}\\LocalServer32
        = D:\\...\\WPS Office\\...\\office6\\wpp.exe /Automation
此时 Dispatch("PowerPoint.Application") 拿到的是 WPS 的一个“僵尸”对象：
Application 级别的调用（Presentations.Count / Add）能过，但拿回来的
Presentation 对象每一次成员调用都会被拒绝（RPC_E_CALL_REJECTED），
pywin32 于是抛出 AttributeError("Open.Slides") —— 用户看到的就是
弹窗里那句莫名其妙的 "Open.Slides"。

v2 的做法
---------
1. 自己找到真正的 POWERPNT.EXE（注册表 App Paths / InstallRoot / 目录扫描）；
2. 写一个内置的空白锚点 pptx，用它启动 PowerPoint；
3. 在运行对象表（ROT）里按锚点文件路径找回**这个** PowerPoint 实例，
   直接取到 Presentation -> Application，绕开被劫持的 ProgID；
4. WithWindow=False 隐藏打开目标文件 → 逐页导出 PNG → 拼回新 PPT；
5. 若本机没有正版 PowerPoint，则退回旧的 Dispatch 方案（并在失败时给出
   能看懂的中文提示）。

对外接口保持不变：
    flatten_ppt(ppt_path, out_path, dpi=192, progress_cb=None) -> out_path
"""
import base64
import gzip
import os
import shutil
import subprocess
import tempfile
import time
import traceback

import paths

_PP_CLSID = "{91493441-5A91-11CF-8700-00AA0060263B}"
_PP_LAYOUT_BLANK = 12          # ppLayoutBlank
_PP_SAVE_AS_PPTX = 24          # ppSaveAsOpenXMLPresentation
_RPC_E_CALL_REJECTED = -2147418111


# --------------------------------------------------------------------------- #
# 小工具
# --------------------------------------------------------------------------- #
def _log(msg):
    """PPTDEBUG=1 时把过程写进 %TEMP%\\ppt_slide.log，方便排错。"""
    if os.environ.get("PPTDEBUG") != "1":
        return
    try:
        p = paths.log_file("ppt_slide.log") or os.path.join(
            tempfile.gettempdir(), "ppt_slide.log")
        with paths.open_binary(p, "a", encoding="utf-8") as f:
            f.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), msg))
    except Exception:
        pass


def _candidate_ok(path):
    try:
        return bool(path) and paths.is_file(path)
    except Exception:
        return False


def _com_path(p):
    """给 PowerPoint COM 用的路径。

    COM/Office 内部仍有 260 字符限制，且不认 \\\\?\\ 前缀：路径过长时换 8.3 短名。
    """
    p = paths.normalize(p)
    if len(p) < 240:
        return p
    short = paths.short_path(p)
    if short and len(short) < len(p):
        return short
    return p


def _office_roots():
    """所有可能装 Office 的 Program Files 根目录（盘符/目录名都不写死）。"""
    roots = []

    def add(p):
        if p and paths.is_dir(p) and p not in roots:
            roots.append(p)

    for env in ("ProgramW6432", "ProgramFiles", "ProgramFiles(x86)"):
        add(os.environ.get(env))
    # Office 完全可能装在 D:/E: 等非系统盘
    try:
        import string
        for letter in string.ascii_uppercase[2:]:
            drive = letter + ":\\"
            if not paths.is_dir(drive):
                continue
            for sub in ("Program Files", "Program Files (x86)",
                        "ProgramFiles", "Office", "Microsoft Office"):
                add(os.path.join(drive, sub))
    except Exception:
        pass
    return roots


def _reg_read(root, subkey, name=""):
    """读注册表值，64/32 位视图都试；失败返回 None。"""
    try:
        import winreg
    except Exception:
        return None
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        try:
            with winreg.OpenKey(root, subkey, 0, winreg.KEY_READ | view) as key:
                val = winreg.QueryValueEx(key, name)[0]
                if val:
                    return str(val)
        except Exception:
            continue
    return None


def find_powerpoint_exe():
    """返回 PowerPoint 的路径；找不到返回 None。

    Windows 返回正版 POWERPNT.EXE 的完整路径；
    macOS 返回 Microsoft PowerPoint.app 的 bundle 路径。
    """
    if paths.MACOS:
        for p in ("/Applications/Microsoft PowerPoint.app",
                  os.path.expanduser("~/Applications/Microsoft PowerPoint.app")):
            if paths.is_dir(p):
                return p
        try:
            out = subprocess.run(
                ["mdfind", "kMDItemCFBundleIdentifier == 'com.microsoft.Powerpoint'"],
                capture_output=True, timeout=8).stdout.decode("utf-8", "ignore")
            for line in out.splitlines():
                line = line.strip()
                if line.endswith(".app") and paths.is_dir(line):
                    return line
        except Exception:
            pass
        return None

    try:
        import winreg
    except Exception:
        return None

    # 1) App Paths（最可靠，用户手动指定过也能找到）
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        val = _reg_read(
            hive,
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE")
        if val:
            val = paths.resolve_existing(val.strip().strip('"'))
            if _candidate_ok(val):
                return val

    # 2) InstallRoot（按版本）
    for ver in ("16.0", "15.0", "14.0", "12.0"):
        root = _reg_read(
            winreg.HKEY_LOCAL_MACHINE,
            r"SOFTWARE\Microsoft\Office\%s\PowerPoint\InstallRoot" % ver, "Path")
        if root:
            cand = os.path.join(root, "POWERPNT.EXE")
            if _candidate_ok(cand):
                return cand

    # 3) ClickToRun 安装位置（Microsoft 365 / 2016+ 的实际布局）
    for sub in (r"SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                r"SOFTWARE\Microsoft\Office\ClickToRun\PropertyBag"):
        for name in ("InstallPath", "PackageRoot", "LastScenarioInstallPath"):
            base = _reg_read(winreg.HKEY_LOCAL_MACHINE, sub, name)
            if not base:
                continue
            base = base.strip().strip('"')
            for rel in ("POWERPNT.EXE", os.path.join("root", "Office16", "POWERPNT.EXE"),
                        os.path.join("Office16", "POWERPNT.EXE")):
                cand = os.path.join(base, rel)
                if _candidate_ok(cand):
                    return cand

    # 4) 目录扫描（不写死盘符和安装位置）
    import glob
    pats = []
    for root in _office_roots():
        pats.extend([
            os.path.join(root, "Microsoft Office", "root", "Office*", "POWERPNT.EXE"),
            os.path.join(root, "Microsoft Office", "Office*", "POWERPNT.EXE"),
            os.path.join(root, "Microsoft Office*", "root", "Office*", "POWERPNT.EXE"),
            os.path.join(root, "Microsoft Office*", "POWERPNT.EXE"),
            os.path.join(root, "Office*", "POWERPNT.EXE"),
            os.path.join(root, "root", "Office*", "POWERPNT.EXE"),
        ])
    for pat in pats:
        try:
            hits = sorted(glob.glob(pat), reverse=True)
        except Exception:
            hits = []
        for hit in hits:
            if _candidate_ok(hit):
                return hit

    # 5) 最后看一眼 PowerPoint.Application 的 LocalServer32（WPS 接管时指向 wpp.exe，会被过滤掉）
    server = _reg_read(
        winreg.HKEY_CLASSES_ROOT,
        r"CLSID\%s\LocalServer32" % _PP_CLSID)
    if server:
        if server.startswith('"'):
            cand = server[1:].split('"', 1)[0]
        else:
            low = server.lower()
            cut = low.find(".exe")
            cand = server[:cut + 4] if cut >= 0 else server.strip()
        if os.path.basename(cand).lower() == "powerpnt.exe" and _candidate_ok(cand):
            return cand
    return None


def _ms_ppt_running():
    """此刻是否已经有正版 POWERPNT.EXE 在运行（决定干完活要不要替用户关掉）。

    不能用注册表/ROT 判断：WPS 接管了同一个 CLSID 后也会注册类对象，
    必须按进程名判断。
    """
    try:
        import ctypes
        from ctypes import wintypes

        TH32CS_SNAPPROCESS = 0x00000002
        INVALID_HANDLE_VALUE = ctypes.c_void_p(-1).value

        class PROCESSENTRY32W(ctypes.Structure):
            _fields_ = [
                ("dwSize", wintypes.DWORD),
                ("cntUsage", wintypes.DWORD),
                ("th32ProcessID", wintypes.DWORD),
                ("th32DefaultHeapID", ctypes.c_size_t),
                ("th32ModuleID", wintypes.DWORD),
                ("cntThreads", wintypes.DWORD),
                ("th32ParentProcessID", wintypes.DWORD),
                ("pcPriClassBase", ctypes.c_long),
                ("dwFlags", wintypes.DWORD),
                ("szExeFile", wintypes.WCHAR * 260),
            ]

        k32 = ctypes.WinDLL("kernel32", use_last_error=True)
        k32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
        k32.CreateToolhelp32Snapshot.argtypes = [wintypes.DWORD, wintypes.DWORD]
        k32.Process32FirstW.argtypes = [wintypes.HANDLE,
                                        ctypes.POINTER(PROCESSENTRY32W)]
        k32.Process32NextW.argtypes = [wintypes.HANDLE,
                                       ctypes.POINTER(PROCESSENTRY32W)]
        k32.CloseHandle.argtypes = [wintypes.HANDLE]

        snap = k32.CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
        if not snap or snap == INVALID_HANDLE_VALUE:
            return False
        try:
            entry = PROCESSENTRY32W()
            entry.dwSize = ctypes.sizeof(PROCESSENTRY32W)
            ok = k32.Process32FirstW(snap, ctypes.byref(entry))
            while ok:
                if str(entry.szExeFile).lower() == "powerpnt.exe":
                    return True
                ok = k32.Process32NextW(snap, ctypes.byref(entry))
        finally:
            k32.CloseHandle(snap)
    except Exception:
        pass
    return False


def _write_anchor(folder):
    """把内置的空白 pptx 释放到磁盘，作为启动 PowerPoint 的锚点。"""
    from anchor_data import ANCHOR_B64
    data = gzip.decompress(base64.b64decode(ANCHOR_B64.encode("ascii")))
    paths.ensure_dir(folder)
    path = os.path.join(folder, "pptfx_anchor_%d.pptx" % int(time.time() * 1000))
    with paths.open_binary(path, "wb") as f:
        f.write(data)
    return path


def _rot_find(path, timeout=150.0, tick_cb=None):
    """在 ROT 里按文件路径找到 PowerPoint 打开的那个演示文稿 moniker。"""
    import pythoncom

    target = os.path.normcase(paths.abspath(path))
    names = {target}
    for alt in (paths.short_path(path), paths.long_name_path(path),
                paths.resolve_existing(path)):
        if alt:
            names.add(os.path.normcase(paths.abspath(alt)))
    base = os.path.normcase(os.path.basename(target))
    deadline = time.time() + timeout
    last_tick = 0.0
    while time.time() < deadline:
        try:
            rot = pythoncom.GetRunningObjectTable()
            ctx = pythoncom.CreateBindCtx(0)
            for mon in rot.EnumRunning():
                try:
                    name = mon.GetDisplayName(ctx, None)
                except Exception:
                    continue
                if not name:
                    continue
                low = os.path.normcase(paths.normalize(name))
                if low in names or os.path.basename(low) == base:
                    return mon, rot
        except Exception:
            pass
        now = time.time()
        if tick_cb and now - last_tick > 2.0:
            last_tick = now
            tick_cb()
        time.sleep(0.25)
    return None, None


def _bind_from_moniker(mon, rot):
    """moniker -> Presentation 对象（同时校验确实是微软 PowerPoint）。"""
    import pythoncom
    import win32com.client

    unknown = rot.GetObject(mon)
    disp = unknown.QueryInterface(pythoncom.IID_IDispatch)
    pres = win32com.client.Dispatch(disp)
    app = pres.Application
    where = ""
    try:
        where = str(app.Path)
    except Exception:
        pass
    _log("attached app.Path=%s" % where)
    lowered = where.lower()
    if ("kingsoft" in lowered) or ("wps" in lowered):
        raise RuntimeError(
            "当前 PowerPoint 接口被 WPS 接管了（%s）。\n"
            "请先退出 WPS，或安装 Microsoft PowerPoint 后重试。" % where)
    return pres, app


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def flatten_ppt(ppt_path, out_path, dpi=192, progress_cb=None):
    """把 PPT 每页导出为图片再拼回新 PPT。源文件只读，输出到 out_path。"""
    def cb(msg):
        if progress_cb:
            try:
                progress_cb(msg)
            except Exception:
                pass

    ppt_path = paths.resolve_existing(ppt_path)
    out_path = paths.abspath(out_path)
    if not paths.is_file(ppt_path):
        reasons = "\n".join("· " + r for r in paths.describe_missing(ppt_path))
        raise FileNotFoundError("找不到源文件：\n%s\n%s" % (ppt_path, reasons))

    if paths.MACOS:
        return _flatten_mac(ppt_path, out_path, dpi, cb)

    exe = find_powerpoint_exe()
    _log("powerpoint exe = %r" % exe)
    if not exe:
        # 没有正版 PowerPoint：退回旧的 COM 方案（WPS 或其它兼容实现）
        return _flatten_via_dispatch(ppt_path, out_path, dpi, cb)

    import pythoncom
    try:
        pythoncom.CoInitialize()
        _co_inited = True
    except Exception:
        _co_inited = False

    tmp_dir = paths.temp_dir("pptfx_")
    app = None
    owned = False
    src_pres = None
    new_pres = None
    try:
        cb("正在启动 PowerPoint 引擎…")
        ext = os.path.splitext(ppt_path)[1] or ".pptx"
        src_copy = os.path.join(tmp_dir, "source%s" % ext)
        shutil.copy2(paths.long_path(ppt_path), paths.long_path(src_copy))

        anchor = _write_anchor(tmp_dir)
        preexisting = _ms_ppt_running()
        try:
            subprocess.Popen([exe, anchor], close_fds=True,
                             stdout=subprocess.DEVNULL,
                             stderr=subprocess.DEVNULL)
        except Exception as exc:
            raise RuntimeError("启动 PowerPoint 失败：%s" % exc)

        mon, rot = _rot_find(anchor, 150.0,
                             tick_cb=lambda: cb("正在等待 PowerPoint 启动…"))
        if mon is None:
            raise RuntimeError(
                "PowerPoint 启动超时（150 秒）。\n\n"
                "可能原因：\n"
                "1. 首次启动需要激活/登录，请先手动打开一次 PowerPoint；\n"
                "2. 有安全软件拦截了自动化；\n"
                "3. 系统资源紧张，请关闭其他大型程序后重试。")

        anchor_pres, app = _bind_from_moniker(mon, rot)
        owned = not preexisting
        if owned:
            # 自己拉起来的实例：直接隐藏，避免窗口闪烁
            try:
                app.WindowState = 2  # ppWindowMinimized
            except Exception:
                pass
        try:
            anchor_pres.Close()
        except Exception:
            pass

        # ---- 1. 隐藏打开源文件 ----
        cb("正在打开源文件…")
        try:
            src_pres = app.Presentations.Open(src_copy, ReadOnly=True,
                                              Untitled=False,
                                              WithWindow=False)
        except Exception as exc:
            raise RuntimeError("用 PowerPoint 打开源文件失败：%s" % exc)

        slide_count = int(src_pres.Slides.Count)
        if slide_count <= 0:
            raise RuntimeError("源文件里没有任何幻灯片。")
        try:
            sw = float(src_pres.PageSetup.SlideWidth)
            sh = float(src_pres.PageSetup.SlideHeight)
        except Exception:
            sw, sh = 960.0, 540.0
        px_w = max(16, int(round(sw * dpi / 72.0)))
        px_h = max(16, int(round(sh * dpi / 72.0)))
        _log("slides=%d size=%.1fx%.1f px=%dx%d dpi=%s"
             % (slide_count, sw, sh, px_w, px_h, dpi))

        # ---- 2. 逐页导出 ----
        img_dir = os.path.join(tmp_dir, "images")
        paths.ensure_dir(img_dir)
        img_dir_com = _com_path(img_dir)
        images = []
        cb("共 %d 页，开始导出…" % slide_count)
        for i in range(1, slide_count + 1):
            png = os.path.join(img_dir, "slide_%03d.png" % i)
            png_com = os.path.join(img_dir_com, "slide_%03d.png" % i)
            try:
                src_pres.Slides.Item(i).Export(png_com, "PNG", px_w, px_h)
            except Exception as exc:
                raise RuntimeError("导出第 %d 页失败：%s" % (i, exc))
            if not paths.is_file(png):
                raise RuntimeError("第 %d 页没有生成图片，可能页面为空或损坏。" % i)
            images.append(png)
            cb("导出第 %d/%d 页" % (i, slide_count))

        try:
            src_pres.Close()
        except Exception:
            pass
        src_pres = None

        # ---- 3. 拼回新 PPT ----
        cb("正在生成图片化 PPT…")
        new_pres = app.Presentations.Add(0)  # WithWindow=False
        try:
            new_pres.PageSetup.SlideWidth = sw
            new_pres.PageSetup.SlideHeight = sh
        except Exception:
            pass

        total = len(images)
        for idx, img in enumerate(images):
            slide = None
            for layout in (_PP_LAYOUT_BLANK, 7, 1):
                try:
                    slide = new_pres.Slides.Add(idx + 1, layout)
                    break
                except Exception:
                    slide = None
            if slide is None:
                raise RuntimeError("新建第 %d 页失败。" % (idx + 1))
            # 清掉版式自带的占位符，只留整页图片
            try:
                for _ in range(int(slide.Shapes.Count)):
                    slide.Shapes.Item(1).Delete()
            except Exception:
                pass
            try:
                slide.Shapes.AddPicture(
                    FileName=os.path.join(img_dir_com, os.path.basename(img)),
                    LinkToFile=False, SaveWithDocument=True,
                    Left=0, Top=0, Width=sw, Height=sh)
            except Exception as exc:
                raise RuntimeError("插入第 %d 页图片失败：%s" % (idx + 1, exc))
            cb("插入第 %d/%d 页" % (idx + 1, total))

        # 保险：把多出来的页删掉
        try:
            while int(new_pres.Slides.Count) > total:
                new_pres.Slides.Item(int(new_pres.Slides.Count)).Delete()
        except Exception:
            pass

        # ---- 4. 保存 ----
        cb("正在保存文件…")
        _log("save start")
        out_dir = os.path.dirname(out_path)
        if out_dir:
            if not paths.ensure_dir(out_dir):
                raise RuntimeError("无法创建输出目录：\n%s" % out_dir)
        if paths.exists(out_path):
            try:
                os.remove(paths.long_path(out_path))
            except Exception:
                raise RuntimeError(
                    "输出文件被占用，无法覆盖：\n%s\n请关闭同名文件后重试。" % out_path)
        return _save_presentation(new_pres, out_path, tmp_dir)
        try:
            new_pres.Close()
            _log("new pres closed")
        except Exception:
            pass
        new_pres = None
        if not paths.is_file(out_path):
            raise RuntimeError("保存后没有找到输出文件：\n%s" % out_path)
        return out_path

    finally:
        for obj in (src_pres, new_pres):
            try:
                if obj is not None:
                    obj.Close()
            except Exception:
                pass
        if owned and app is not None:
            try:
                _log("quit start")
                app.Quit()
                _log("quit done")
            except Exception as exc:
                _log("quit err %r" % (exc,))
        _log("rmtree start")
        shutil.rmtree(paths.long_path(tmp_dir), ignore_errors=True)
        _log("rmtree done")
        if _co_inited:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass


# --------------------------------------------------------------------------- #
# 兜底方案：直接用 PowerPOint.Application（本机没有正版 PowerPoint 时）
# --------------------------------------------------------------------------- #
def _save_presentation(pres, out_path, tmp_dir):
    """保存到 out_path。

    Office 对 SaveAs 的路径有 255 字符硬限制（超出直接 Invalid request），
    这种情况先存到临时目录，再用文件系统（支持长路径）搬过去。
    """
    primary = _com_path(out_path)
    direct = len(primary) < 250

    if direct:
        try:
            pres.SaveAs(primary, _PP_SAVE_AS_PPTX)
            return out_path
        except Exception as exc:
            _log("saveas primary failed: %r" % (exc,))

    fallback = os.path.join(tmp_dir, "output.pptx")
    try:
        pres.SaveAs(fallback, _PP_SAVE_AS_PPTX)
        _log("saveas fallback ok -> %s" % fallback)
    except Exception as exc:
        _log("saveas fallback failed: %r" % (exc,))
        raise RuntimeError("保存失败：%s" % exc)

    dst = paths.long_path(out_path)
    try:
        shutil.move(fallback, dst)
    except Exception:
        shutil.copy2(fallback, dst)
        try:
            os.remove(fallback)
        except Exception:
            pass
    _log("moved fallback -> %s" % out_path)
    return out_path


def _flatten_via_dispatch(ppt_path, out_path, dpi, progress_cb):
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    src = None
    new_pres = None
    tmp_dir = None
    try:
        try:
            app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as exc:
            raise RuntimeError(
                "本机没有找到可用的 PowerPoint：\n%s\n\n"
                "请安装 Microsoft Office（PowerPoint）后重试。" % exc)

        where = ""
        try:
            where = str(app.Path)
        except Exception:
            pass
        _log("dispatch app.Path=%s" % where)

        # 源文件先复制到临时目录：COM 端对超长路径（>260）几乎必然失败，
        # 而临时目录是可控的短路径。
        tmp_dir = paths.temp_dir("pptfx_")
        ext = os.path.splitext(ppt_path)[1] or ".pptx"
        src_copy = os.path.join(tmp_dir, "source%s" % ext)
        shutil.copy2(paths.long_path(ppt_path), paths.long_path(src_copy))

        try:
            src = app.Presentations.Open(_com_path(src_copy), ReadOnly=True,
                                         Untitled=False, WithWindow=False)
            slide_count = int(src.Slides.Count)
            sw = float(src.PageSetup.SlideWidth)
            sh = float(src.PageSetup.SlideHeight)
        except Exception as exc:
            if "拒绝" in str(exc) or "reject" in str(exc).lower() \
                    or "Slides" in str(exc):
                raise RuntimeError(
                    "PowerPoint 拒绝接受调用，通常是 WPS 正在运行并接管了"
                    " PowerPoint 接口。\n\n请先退出 WPS（右下角托盘也退出），"
                    "或安装 Microsoft PowerPoint 后重试。\n\n技术细节：%s" % exc)
            raise

        px_w = max(16, int(round(sw * dpi / 72.0)))
        px_h = max(16, int(round(sh * dpi / 72.0)))
        try:
            img_dir = os.path.join(tmp_dir, "images")
            paths.ensure_dir(img_dir)
            img_dir_com = _com_path(img_dir)
            images = []
            for i in range(1, slide_count + 1):
                png = os.path.join(img_dir, "slide_%03d.png" % i)
                src.Slides.Item(i).Export(
                    os.path.join(img_dir_com, "slide_%03d.png" % i),
                    "PNG", px_w, px_h)
                images.append(png)
                if progress_cb:
                    progress_cb("导出第 %d/%d 页" % (i, slide_count))
            src.Close()
            src = None

            if progress_cb:
                progress_cb("正在生成图片化 PPT…")
            new_pres = app.Presentations.Add(0)
            try:
                new_pres.PageSetup.SlideWidth = sw
                new_pres.PageSetup.SlideHeight = sh
            except Exception:
                pass
            for idx, img in enumerate(images):
                slide = new_pres.Slides.Add(idx + 1, _PP_LAYOUT_BLANK)
                try:
                    for _ in range(int(slide.Shapes.Count)):
                        slide.Shapes.Item(1).Delete()
                except Exception:
                    pass
                slide.Shapes.AddPicture(
                    FileName=os.path.join(img_dir_com, os.path.basename(img)),
                    LinkToFile=False, SaveWithDocument=True,
                    Left=0, Top=0, Width=sw, Height=sh)
                if progress_cb:
                    progress_cb("插入第 %d/%d 页" % (idx + 1, len(images)))
            out_dir = os.path.dirname(out_path)
            if out_dir:
                paths.ensure_dir(out_dir)
            if paths.exists(out_path):
                try:
                    os.remove(paths.long_path(out_path))
                except Exception:
                    raise RuntimeError("输出文件被占用，无法覆盖：\n%s" % out_path)
            _save_presentation(new_pres, out_path, tmp_dir)
            return out_path
        finally:
            shutil.rmtree(paths.long_path(tmp_dir), ignore_errors=True)
    finally:
        for obj in (src, new_pres):
            try:
                if obj is not None:
                    obj.Close()
            except Exception:
                pass
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        # Open 阶段抛错时，内层 finally 还没接管，这里兜一次
        if tmp_dir:
            shutil.rmtree(paths.long_path(tmp_dir), ignore_errors=True)


# --------------------------------------------------------------------------- #
# macOS 实现
# --------------------------------------------------------------------------- #
_LO_CANDIDATES = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "~/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
)


def find_libreoffice():
    """LibreOffice 的 soffice 路径；找不到返回 None。"""
    for c in _LO_CANDIDATES:
        p = os.path.expanduser(c)
        if _candidate_ok(p):
            return p
    try:
        found = shutil.which("soffice") or shutil.which("libreoffice")
    except Exception:
        found = None
    if found and _candidate_ok(found):
        return found
    return None


def has_pdf_renderer():
    """macOS 的 PDF→PNG 渲染依赖 PyMuPDF。"""
    try:
        import fitz  # noqa: F401
        return True
    except Exception:
        return False


def _mac_export_via_libreoffice(src, tmp_dir, dpi, cb):
    """LibreOffice 转 PDF → PyMuPDF 按 dpi 渲染成 PNG。返回图片路径列表。"""
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("no-libreoffice")
    if not has_pdf_renderer():
        raise RuntimeError("no-pymupdf")

    cb("正在用 LibreOffice 解析页面…")
    work = os.path.join(tmp_dir, "lo")
    paths.ensure_dir(work)
    env = dict(os.environ)
    env["HOME"] = tmp_dir                       # 避免动到用户的 LibreOffice 配置
    cmd = [soffice, "--headless", "--norestore", "--nologo",
           "--convert-to", "pdf", "--outdir", work, src]
    proc = subprocess.run(cmd, env=env, stdout=subprocess.PIPE,
                          stderr=subprocess.STDOUT, timeout=600)
    pdfs = [f for f in os.listdir(work) if f.lower().endswith(".pdf")] \
        if paths.is_dir(work) else []
    if not pdfs:
        tail = (proc.stdout or b"").decode("utf-8", "ignore")[-400:]
        raise RuntimeError("LibreOffice 转换失败：%s" % (tail.strip() or "没有生成 PDF"))

    import fitz
    doc = fitz.open(os.path.join(work, pdfs[0]))
    if doc.page_count <= 0:
        raise RuntimeError("转换后的 PDF 没有任何页面。")
    img_dir = os.path.join(tmp_dir, "images")
    paths.ensure_dir(img_dir)
    zoom = float(dpi) / 72.0
    mat = fitz.Matrix(zoom, zoom)
    images = []
    total = doc.page_count
    for i in range(total):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        png = os.path.join(img_dir, "slide_%03d.png" % (i + 1))
        pix.save(png)
        images.append(png)
        cb("导出第 %d/%d 页" % (i + 1, total))
    doc.close()
    return images


_MAC_EXPORT_SCRIPT = r'''
on run argv
    set inFile to POSIX file (item 1 of argv)
    set outDir to POSIX file (item 2 of argv)
    set dpiArg to (item 3 of argv) as integer
    tell application "Microsoft PowerPoint"
        open inFile
        set thePres to active presentation
        save thePres in outDir as save as PNG
        close thePres saving no
    end tell
    return "ok"
end run
'''


def _mac_export_via_powerpoint(src, tmp_dir, dpi, cb):
    """PowerPoint for Mac 用 AppleScript 导出每页 PNG（首次会请求"自动化"授权）。"""
    if not find_powerpoint_exe():
        raise RuntimeError("no-powerpoint")
    work = os.path.join(tmp_dir, "pp")
    paths.ensure_dir(work)
    cb("正在调用 Microsoft PowerPoint 导出页面…")
    proc = subprocess.run(["/usr/bin/osascript", "-e", _MAC_EXPORT_SCRIPT,
                           src, work, str(int(dpi))],
                          stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                          timeout=900)
    out = (proc.stdout or b"").decode("utf-8", "ignore").strip()
    if proc.returncode != 0:
        raise RuntimeError("PowerPoint 自动化失败：%s" % out[-300:])
    images = []
    for root, _dirs, files in os.walk(work):
        for fn in sorted(files):
            if fn.lower().endswith(".png"):
                images.append(os.path.join(root, fn))
    if not images:
        raise RuntimeError("PowerPoint 没有导出任何 PNG。")
    images.sort()
    for i, p in enumerate(images, 1):
        cb("导出第 %d/%d 页" % (i, len(images)))
    return images


def _pptx_slide_size(p):
    """读源文件的页面尺寸（EMU）；读不到返回 None。"""
    try:
        from pptx import Presentation
        prs = Presentation(p)
        w, h = int(prs.slide_width), int(prs.slide_height)
        if w > 0 and h > 0:
            return w, h
    except Exception:
        pass
    return None


def _build_deck_from_images(images, out_path, dpi, cb, slide_size=None):
    """用 python-pptx 把整页图片拼成新的 pptx。

    页面尺寸优先沿用源文件；读不到就按图片像素换算，保证不变形。
    """
    from pptx import Presentation
    from pptx.util import Emu
    from PIL import Image

    cb("正在生成图片化 PPT…")
    with Image.open(images[0]) as im:
        w_px, h_px = im.size

    if slide_size:
        width, height = slide_size
    else:
        width = 12192000                                   # 13.333 英寸（16:9 常见宽度）
        height = max(1, int(round(width * float(h_px) / max(1, w_px))))

    prs = Presentation()
    prs.slide_width = Emu(width)
    prs.slide_height = Emu(height)
    blank = prs.slide_layouts[6]
    total = len(images)
    for i, img in enumerate(images, 1):
        slide = prs.slides.add_slide(blank)
        slide.shapes.add_picture(img, 0, 0,
                                 width=prs.slide_width, height=prs.slide_height)
        cb("插入第 %d/%d 页" % (i, total))

    out_dir = os.path.dirname(out_path)
    if out_dir and not paths.ensure_dir(out_dir):
        raise RuntimeError("无法创建输出目录：\n%s" % out_dir)
    if paths.exists(out_path):
        try:
            os.remove(paths.long_path(out_path))
        except Exception:
            raise RuntimeError(
                "输出文件被占用，无法覆盖：\n%s\n请关闭同名文件后重试。" % out_path)
    cb("正在保存文件…")
    prs.save(out_path)
    return out_path


def _flatten_mac(ppt_path, out_path, dpi, cb):
    """macOS：导出每页图片 → 拼回新 PPT。

    导出引擎按可用性自动选：LibreOffice（无需授权，优先）→ PowerPoint 自动化。
    """
    tmp_dir = paths.temp_dir("pptfx_mac_")
    try:
        cb("正在准备导出环境…")
        src_copy = os.path.join(tmp_dir, "source" + (os.path.splitext(ppt_path)[1] or ".pptx"))
        shutil.copy2(paths.long_path(ppt_path), paths.long_path(src_copy))

        errors = []
        images = None
        for name, fn in (("libreoffice", _mac_export_via_libreoffice),
                         ("powerpoint", _mac_export_via_powerpoint)):
            try:
                images = fn(src_copy, tmp_dir, dpi, cb)
                _log("mac export via %s ok, %d images" % (name, len(images)))
                break
            except Exception as exc:
                _log("mac export via %s failed: %r" % (name, exc))
                errors.append((name, str(exc)))

        if not images:
            detail = "\n".join("· %s：%s" % (n, e) for n, e in errors)
            if any(e == "no-libreoffice" for _n, e in errors):
                hint = ("本机没有找到可用的导出引擎。\n\n"
                        "图片化需要以下任一环境：\n"
                        "1) 安装 LibreOffice（免费，推荐）：https://zh-cn.libreoffice.org/\n"
                        "2) 安装 Microsoft PowerPoint for Mac\n"
                        "3) 若已装 LibreOffice 仍报错，请确认 PyMuPDF 已随程序打包（自检页可看）")
            else:
                hint = ("导出页面失败。\n\n"
                        "常见原因：\n"
                        "1) 第一次用 PowerPoint 自动化时，系统弹窗里没有点「好」允许控制；\n"
                        "2) 文件加密/损坏，先用 PowerPoint 或 LibreOffice 打开确认；\n"
                        "3) 输出目录没有写权限。")
            raise RuntimeError("页面导出失败：\n%s\n\n%s" % (detail, hint))

        return _build_deck_from_images(images, out_path, dpi, cb,
                                       slide_size=_pptx_slide_size(ppt_path))
    finally:
        shutil.rmtree(paths.long_path(tmp_dir), ignore_errors=True)


if __name__ == "__main__":
    import sys
    src = sys.argv[1] if len(sys.argv) > 1 else None
    print("PowerPoint:", find_powerpoint_exe())
    print("正版实例在跑:", _ms_ppt_running())
    if src:
        out = os.path.splitext(src)[0] + "_图片化.pptx"
        t0 = time.time()
        try:
            flatten_ppt(src, out, dpi=192,
                        progress_cb=lambda m: print("  ", m))
            print("完成 %.1fs -> %s (%d bytes)"
                  % (time.time() - t0, out, os.path.getsize(out)))
        except Exception:
            traceback.print_exc()
