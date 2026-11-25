# -*- coding: utf-8 -*-
"""
自动保存 / 定时备份引擎（不依赖 tkinter，可单独测试）

设计要点：
1. 只附着到**已经在运行**的 PowerPoint 或 WPS 演示进程，绝不新建进程，
   也不会去打开/关闭用户的文档。
2. 所有 COM 调用都在一次性短线程里完成并带超时：PowerPoint 弹着对话框
   卡住时，主循环不会被拖死，只记一条“无响应，本次跳过”。
3. 兼容链路：GetActiveObject(PowerPoint.Application / KWPP.Application /
   WPP.Application) → 失败再用 ROT 枚举已打开的演示文稿兜底。
4. 备份用 SaveCopyAs（另存副本），不改动原文件、不产生“未保存”状态；
   自动保存用 Save()，只在文档确实有改动（Saved == False）时调用。
"""
import os
import re
import time
import json
import shutil
import threading
from collections import deque

import paths

# --------------------------------------------------------------------------- #
# 配置
# --------------------------------------------------------------------------- #
DEFAULTS = {
    "enabled": False,          # 上次退出时是否处于监听状态
    "save_enabled": True,      # 编辑时自动保存
    "save_interval_min": 10,
    "backup_enabled": True,    # 定时备份到指定目录
    "backup_interval_min": 30,
    "backup_dir": "",
    "keep_copies": 60,         # 每个文档最多保留多少份备份
}

PPT_EXTS = (".pptx", ".pptm", ".ppt", ".pps", ".ppsx", ".ppsm",
            ".pot", ".potx", ".potm")

# SaveCopyAs 的 FileFormat：按源格式选，避免 .pptm 被存成无宏版本
FMT_MAP = {
    ".pptx": 24,   # ppSaveAsOpenXMLPresentation
    ".pptm": 25,   # ppSaveAsOpenXMLPresentationMacroEnabled
    ".ppsx": 28,   # ppSaveAsOpenXMLShow
    ".ppsm": 29,   # ppSaveAsOpenXMLShowMacroEnabled
    ".potx": 26,   # ppSaveAsOpenXMLTemplate
    ".potm": 27,   # ppSaveAsOpenXMLTemplateMacroEnabled
    ".ppt": 1,     # ppSaveAsPresentation
    ".pps": 1,
    ".pot": 1,
}


def config_path():
    return paths.cache_file("autosave.json")


def load_config():
    cfg = dict(DEFAULTS)
    try:
        raw = paths.read_text(config_path())
        data = json.loads(raw)
        if isinstance(data, dict):
            for k, v in data.items():
                if k in DEFAULTS:
                    cfg[k] = v
    except Exception:
        pass
    cfg["save_interval_min"] = clamp_interval(cfg.get("save_interval_min"), 10)
    cfg["backup_interval_min"] = clamp_interval(cfg.get("backup_interval_min"), 30)
    try:
        cfg["keep_copies"] = max(1, min(500, int(cfg.get("keep_copies") or 60)))
    except Exception:
        cfg["keep_copies"] = 60
    if not isinstance(cfg.get("backup_dir"), str):
        cfg["backup_dir"] = ""
    return cfg


def save_config(cfg):
    try:
        paths.write_text(config_path(), json.dumps(cfg, ensure_ascii=False, indent=2))
        return True
    except Exception:
        return False


def clamp_interval(value, default=10):
    """间隔允许 1~720 分钟；也接受 0.1 分钟（6 秒）便于自测。"""
    try:
        v = float(value)
    except Exception:
        return default
    if v <= 0:
        return default
    return max(0.1, min(720.0, v))


def default_backup_dir():
    doc = paths.known_folder(getattr(paths, "FOLDERID_DOCUMENTS",
                                      "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"))
    base = doc or paths.user_data_dir()
    return os.path.join(base, "PPT自动备份")


def interval_text(minutes):
    try:
        m = float(minutes)
    except Exception:
        return "—"
    if m < 1:
        return "%d 秒" % int(round(m * 60))
    if abs(m - round(m)) < 0.01:
        return "%d 分钟" % int(round(m))
    return "%s 分钟" % ("%.1f" % m)


# --------------------------------------------------------------------------- #
# COM：附着到正在运行的演示文稿
# --------------------------------------------------------------------------- #
def _safe(fn, default=""):
    try:
        return fn()
    except Exception:
        return default


def _app_label(app):
    """返回 (友好名, 原始 Name, 安装路径)"""
    name = str(_safe(lambda: app.Name, ""))
    where = str(_safe(lambda: app.Path, ""))
    low = (name + " " + where).lower()
    if "kingsoft" in low or "wps" in low:
        return "WPS 演示", name, where
    if "powerpoint" in low:
        return "Microsoft PowerPoint", name, where
    return (name or "演示文稿程序"), name, where


def _apps():
    """把所有能附着上的演示程序都找出来（PowerPoint 与 WPS 可以同时运行）。

    返回 [(友好名, app)]；一个都没有时抛 RuntimeError。
    macOS 版只支持 Microsoft PowerPoint（WPS for Mac 不提供脚本接口）。
    """
    if paths.MACOS:
        if _mac_running():
            return [("Microsoft PowerPoint", "mac")]
        raise RuntimeError("未检测到正在运行的 Microsoft PowerPoint"
                           "（macOS 版暂不支持 WPS 演示）")

    import win32com.client

    errors = []
    out = []
    seen = set()
    for progid in ("PowerPoint.Application", "KWPP.Application", "WPP.Application"):
        try:
            app = win32com.client.GetActiveObject(progid)
        except Exception as exc:
            errors.append("%s: %s" % (progid, exc.__class__.__name__))
            continue
        label, name, where = _app_label(app)
        key = (label.lower(), os.path.normcase(where or name))
        if key in seen:
            continue
        seen.add(key)
        out.append((label, app))

    if out:
        return out

    # 兜底：ROT 里找任意一个已打开的演示文稿，顺着它拿 Application
    try:
        import pythoncom
        app = _attach_via_rot(pythoncom, win32com.client)
        if app is not None:
            label, _, _ = _app_label(app)
            return [(label, app)]
    except Exception as exc:
        errors.append("ROT: %s" % exc.__class__.__name__)

    raise RuntimeError("未检测到正在运行的 PowerPoint 或 WPS 演示（%s）"
                       % "；".join(errors[:2]))


def _attach():
    """兼容旧调用：只返回第一个应用。"""
    apps = _apps()
    return apps[0][1], apps[0][0]


def _attach_via_rot(pythoncom, win32com_client):
    rot = pythoncom.GetRunningObjectTable()
    ctx = pythoncom.CreateBindCtx(0)
    for mon in rot.EnumRunning():
        try:
            name = mon.GetDisplayName(ctx, None) or ""
        except Exception:
            continue
        if not name.lower().endswith(PPT_EXTS):
            continue
        try:
            unknown = rot.GetObject(mon)
            disp = unknown.QueryInterface(pythoncom.IID_IDispatch)
            pres = win32com_client.Dispatch(disp)
            app = pres.Application
            if int(app.Presentations.Count) >= 1:
                return app
        except Exception:
            continue
    return None


def _docs(app, label=""):
    """列出某个应用里打开的所有演示文稿（纯数据 + 对象引用）"""
    if paths.MACOS:
        return _mac_docs(label)
    out = []
    coll = app.Presentations
    try:
        count = int(coll.Count)
    except Exception:
        return out
    for i in range(1, count + 1):
        try:
            pres = coll.Item(i)
        except Exception:
            continue
        out.append({
            "pres": pres,
            "app": label,
            "name": str(_safe(lambda: pres.Name, "未命名")),
            "full": str(_safe(lambda: pres.FullName, "") or ""),
            "saved": bool(_safe(lambda: pres.Saved, True)),
            "readonly": bool(_safe(lambda: pres.ReadOnly, False)),
        })
    return out


def _all_docs(apps):
    """合并多个应用里的文档，按 程序+路径 去重。"""
    docs = []
    seen = set()
    for label, app in apps:
        for d in _docs(app, label):
            key = (label, os.path.normcase(d["full"]) if d["full"]
                   else ("?" + d["name"]))
            if key in seen:
                continue
            seen.add(key)
            docs.append(d)
    return docs


# --------------------------------------------------------------------------- #
# macOS：AppleScript 后端（只支持 Microsoft PowerPoint）
# --------------------------------------------------------------------------- #
_MAC_LIST_SCRIPT = r'''
set out to ""
tell application "Microsoft PowerPoint"
    if (count of presentations) is 0 then return ""
    repeat with p in presentations
        set nm to ""
        set fp to ""
        set sv to true
        set ro to false
        try
            set nm to (name of p) as text
        end try
        try
            set fp to (full name of p) as text
            if fp is not "" then
                try
                    set fp to POSIX path of (fp as alias)
                end try
            end if
        end try
        try
            set sv to saved of p
        end try
        try
            set ro to read only of p
        end try
        set out to out & nm & "\t" & fp & "\t" & (sv as text) & "\t" & (ro as text) & "\n"
    end repeat
end tell
return out
'''

_MAC_SAVE_SCRIPT = r'''
tell application "Microsoft PowerPoint"
    set did to 0
    repeat with p in presentations
        try
            if (full name of p is not "") and (saved of p is false) then
                save p
                set did to did + 1
            end if
        end try
    end repeat
    return did
end tell
'''


def _mac_running():
    """PowerPoint 进程是否在跑（不启动、不激活窗口）。"""
    try:
        import subprocess
        r = subprocess.run(["pgrep", "-x", "Microsoft PowerPoint"],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           timeout=5)
        return r.returncode == 0
    except Exception:
        return False


def _osascript(script, timeout=25):
    """跑一段 AppleScript，返回 stdout 文本；失败抛 RuntimeError。"""
    import subprocess
    try:
        proc = subprocess.run(["/usr/bin/osascript", "-e", script],
                              stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                              timeout=timeout)
    except subprocess.TimeoutExpired:
        raise RuntimeError("PowerPoint 无响应（脚本超时），本次跳过")
    if proc.returncode != 0:
        err = (proc.stderr or b"").decode("utf-8", "ignore").strip()
        raise RuntimeError(err.splitlines()[-1] if err else "AppleScript 执行失败")
    return (proc.stdout or b"").decode("utf-8", "ignore")


def _mac_docs(label="Microsoft PowerPoint"):
    out = []
    try:
        text = _osascript(_MAC_LIST_SCRIPT)
    except Exception:
        return out
    for line in text.splitlines():
        parts = line.split("\t")
        if not parts or not parts[0].strip():
            continue
        name = parts[0].strip()
        full = parts[1].strip() if len(parts) > 1 else ""
        saved = (parts[2].strip().lower() != "false") if len(parts) > 2 else True
        ro = (parts[3].strip().lower() == "true") if len(parts) > 3 else False
        out.append({"pres": None, "app": label, "name": name,
                    "full": full, "saved": saved, "readonly": ro})
    return out


def _work_mac(cfg, want_save, want_backup):
    """macOS 版一次「保存 → 备份」：保存走 AppleScript，备份直接复制磁盘文件。

    备份复制的是保存之后磁盘上的文件，不碰用户文档、不改变文档路径。
    """
    result = {"app": "", "docs": [], "saved": [], "backups": [], "notes": [],
              "errors": [], "removed": 0, "readonly": []}
    if not _mac_running():
        result["errors"].append("未检测到正在运行的 Microsoft PowerPoint")
        return result
    result["app"] = "Microsoft PowerPoint"

    try:
        docs = _mac_docs()
    except Exception as exc:
        result["errors"].append(str(exc))
        return result
    result["docs"] = [{"name": d["name"], "full": d["full"], "saved": d["saved"],
                       "readonly": d["readonly"], "app": d["app"]} for d in docs]
    if not docs:
        result["notes"].append("程序已连接，但没有打开的演示文稿")
        return result

    if want_save:
        try:
            _osascript(_MAC_SAVE_SCRIPT, timeout=40)
            for d in docs:
                if d["full"] and not d["saved"] and not d["readonly"]:
                    result["saved"].append("%s（%s）" % (d["name"], d["app"]))
                elif not d["full"]:
                    result["notes"].append(
                        "%s：新建后还没保存过，请先按 ⌘S 命名一次" % d["name"])
        except Exception as exc:
            result["errors"].append("自动保存失败：%s" % exc)

    if want_backup:
        folder = cfg.get("backup_dir") or default_backup_dir()
        try:
            paths.ensure_dir(folder)
        except Exception as exc:
            result["errors"].append("备份目录不可用：%s" % exc)
            return result
        if not paths.is_dir(folder):
            result["errors"].append("备份目录不存在：%s" % folder)
            return result
        keep = int(cfg.get("keep_copies") or 60)
        stamp = time.strftime("%Y%m%d_%H%M%S")
        for d in docs:
            try:
                if not d["full"]:
                    result["notes"].append("%s：新建后还没保存过，无法备份" % d["name"])
                    continue
                if not paths.is_file(d["full"]):
                    result["notes"].append("%s：磁盘上找不到文件，已跳过" % d["name"])
                    continue
                ext = os.path.splitext(d["full"])[1].lower()
                if ext not in PPT_EXTS:
                    continue
                stem = os.path.splitext(os.path.basename(d["full"]))[0]
                dest = _unique_path(os.path.join(folder, "%s_%s%s" % (stem, stamp, ext)))
                shutil.copy2(paths.long_path(d["full"]), paths.long_path(dest))
                result["backups"].append(os.path.basename(dest))
                result["removed"] += _prune(folder, stem, ext, keep)
            except Exception as exc:
                result["errors"].append("备份 %s 失败：%s" % (d["name"], exc))
    return result


# --------------------------------------------------------------------------- #
# 具体动作
# --------------------------------------------------------------------------- #
def _save_copy(pres, dest, ext):
    """SaveCopyAs；超长路径退化为「先存临时目录再搬过去」。"""
    fmt = FMT_MAP.get(ext)
    last = None
    for kwargs in ((dest, fmt), (dest,)):
        if kwargs[-1] is None:
            continue
        try:
            if len(kwargs) == 2:
                pres.SaveCopyAs(kwargs[0], kwargs[1])
            else:
                pres.SaveCopyAs(kwargs[0])
            return dest
        except Exception as exc:
            last = exc
    # 长路径 / 中文路径：先落到短临时目录
    tmp_dir = paths.temp_dir("asbk_")
    tmp = os.path.join(tmp_dir, "copy" + ext)
    try:
        try:
            pres.SaveCopyAs(tmp, fmt)
        except Exception:
            pres.SaveCopyAs(tmp)
        paths.ensure_dir(os.path.dirname(dest))
        shutil.move(tmp, paths.long_path(dest))
        return dest
    except Exception:
        raise last or RuntimeError("另存副本失败")
    finally:
        shutil.rmtree(paths.long_path(tmp_dir), ignore_errors=True)


def _unique_path(dest):
    if not paths.exists(dest):
        return dest
    stem, ext = os.path.splitext(dest)
    for i in range(1, 100):
        cand = "%s_%d%s" % (stem, i, ext)
        if not paths.exists(cand):
            return cand
    return dest


def _prune(folder, stem, ext, keep):
    """每个文档只保留最近 keep 份备份。"""
    try:
        names = os.listdir(paths.long_path(folder))
    except Exception:
        return 0
    pat = re.compile(r"^" + re.escape(stem) + r"_\d{8}_\d{6}(_\d+)?"
                     + re.escape(ext) + r"$", re.IGNORECASE)
    files = [n for n in names if pat.match(n)]
    if len(files) <= keep:
        return 0
    files.sort()  # 时间戳在文件名里，字典序即时间序
    removed = 0
    for name in files[:len(files) - keep]:
        try:
            os.remove(paths.long_path(os.path.join(folder, name)))
            removed += 1
        except Exception:
            pass
    return removed


def _work(cfg, want_save, want_backup):
    """在独立线程里跑完一次「附着 → 保存 → 备份」，只回传纯数据。"""
    if paths.MACOS:
        return _work_mac(cfg, want_save, want_backup)

    import pythoncom

    pythoncom.CoInitialize()
    result = {"app": "", "docs": [], "saved": [], "backups": [], "notes": [],
              "errors": [], "removed": 0, "readonly": []}
    try:
        try:
            apps = _apps()
        except Exception as exc:
            result["errors"].append(str(exc))
            return result
        result["app"] = " + ".join(label for label, _ in apps)

        docs = _all_docs(apps)
        result["docs"] = [{"name": d["name"], "full": d["full"],
                           "saved": d["saved"], "readonly": d["readonly"],
                           "app": d["app"]}
                          for d in docs]
        if not docs:
            result["notes"].append("程序已连接，但没有打开的演示文稿")
            return result

        if want_save:
            for d in docs:
                try:
                    if d["readonly"]:
                        result["notes"].append("%s：只读，已跳过" % d["name"])
                        continue
                    if not d["full"]:
                        result["notes"].append(
                            "%s：新建后还没保存过，请先按 Ctrl+S 命名一次" % d["name"])
                        continue
                    if d["saved"]:
                        continue
                    d["pres"].Save()
                    result["saved"].append("%s（%s）" % (d["name"], d["app"]))
                except Exception as exc:
                    result["errors"].append("保存 %s 失败：%s" % (d["name"], exc))

        if want_backup:
            folder = cfg.get("backup_dir") or default_backup_dir()
            try:
                paths.ensure_dir(folder)
            except Exception as exc:
                result["errors"].append("备份目录不可用：%s" % exc)
                return result
            if not paths.is_dir(folder):
                result["errors"].append("备份目录不存在：%s" % folder)
                return result
            keep = int(cfg.get("keep_copies") or 60)
            stamp = time.strftime("%Y%m%d_%H%M%S")
            for d in docs:
                try:
                    if not d["full"]:
                        result["notes"].append(
                            "%s：新建后还没保存过，无法备份" % d["name"])
                        continue
                    ext = os.path.splitext(d["full"])[1].lower()
                    if ext not in PPT_EXTS:
                        continue
                    stem = os.path.splitext(os.path.basename(d["full"]))[0]
                    dest = _unique_path(os.path.join(folder, "%s_%s%s" % (stem, stamp, ext)))
                    _save_copy(d["pres"], dest, ext)
                    result["backups"].append(os.path.basename(dest))
                    result["removed"] += _prune(folder, stem, ext, keep)
                except Exception as exc:
                    result["errors"].append("备份 %s 失败：%s" % (d["name"], exc))
        return result
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


# --------------------------------------------------------------------------- #
# 引擎
# --------------------------------------------------------------------------- #
class Engine:
    """后台监听线程。UI 只读 snapshot() / 调 start()、stop()、force_backup()。"""

    def __init__(self, cfg_provider):
        self._cfg_provider = cfg_provider
        self._lock = threading.Lock()
        self._stop_evt = threading.Event()
        self._thread = None
        self._logs = deque(maxlen=80)
        self._next_save = 0.0
        self._next_backup = 0.0
        self._force_backup = False
        self._status_at = 0.0
        self._last_sig = None
        self.status = {
            "running": False, "attached": False, "app": "", "docs": 0,
            "doc_names": [], "last_save": 0.0, "last_backup": 0.0,
            "error": "", "busy": False, "updated": 0.0,
        }

    # ---------------- 生命周期 ----------------
    @property
    def running(self):
        return bool(self._thread and self._thread.is_alive())

    def start(self):
        if self.running:
            return True
        self._stop_evt.clear()
        now = time.time()
        cfg = self._cfg_provider()
        self._next_save = now + clamp_interval(cfg.get("save_interval_min")) * 60
        self._next_backup = now + clamp_interval(cfg.get("backup_interval_min")) * 60
        self._thread = threading.Thread(target=self._loop, name="pptfx-autosave",
                                        daemon=True)
        self._thread.start()
        with self._lock:
            self.status["running"] = True
        self.log("info", "开始监听（自动保存每 %s，备份每 %s）" % (
            interval_text(cfg.get("save_interval_min")),
            interval_text(cfg.get("backup_interval_min"))))
        return True

    def stop(self):
        self._stop_evt.set()
        t = self._thread
        if t:
            t.join(timeout=3.0)
        self._thread = None
        with self._lock:
            self.status["running"] = False
            self.status["attached"] = False
            self.status["busy"] = False
        self.log("info", "已停止监听")

    def force_backup(self):
        self._force_backup = True
        self.log("info", "已请求立即备份，正在执行…")
        return True

    def reset_timers(self):
        """改过间隔后重新计时，避免沿用旧的到期时间。"""
        now = time.time()
        cfg = self._cfg_provider()
        self._next_save = now + clamp_interval(cfg.get("save_interval_min")) * 60
        self._next_backup = now + clamp_interval(cfg.get("backup_interval_min")) * 60

    def next_save_at(self):
        return self._next_save if self.running else 0.0

    def next_backup_at(self):
        return self._next_backup if self.running else 0.0

    # ---------------- 日志 / 状态 ----------------
    def log(self, kind, msg):
        with self._lock:
            self._logs.append((time.time(), kind, msg))

    def snapshot(self):
        with self._lock:
            st = dict(self.status)
            st["doc_names"] = list(st.get("doc_names") or [])
            logs = list(self._logs)
        st["running"] = self.running
        return st, logs

    def _update(self, **kw):
        with self._lock:
            self.status.update(kw)
            self.status["updated"] = time.time()

    # ---------------- 主循环 ----------------
    def _loop(self):
        try:
            import pythoncom
            pythoncom.CoInitialize()
        except Exception:
            pass
        try:
            while not self._stop_evt.is_set():
                try:
                    self._tick()
                except Exception as exc:
                    self.log("error", "内部错误：%s" % exc)
                self._stop_evt.wait(1.0)
        finally:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _tick(self):
        cfg = self._cfg_provider()
        now = time.time()
        save_on = bool(cfg.get("save_enabled"))
        backup_on = bool(cfg.get("backup_enabled"))
        due_save = save_on and now >= self._next_save
        due_backup = (backup_on and now >= self._next_backup) or self._force_backup
        need_status = (now - self._status_at) >= 3.0

        if not (save_on or backup_on):
            if need_status:
                self._status_at = now
                self._update(attached=False, app="", docs=0, doc_names=[],
                             error="")
            return
        if not (due_save or due_backup or need_status):
            return

        # 先推进计时器：即使这次失败也不会疯狂重试
        if due_save:
            self._next_save = now + clamp_interval(cfg.get("save_interval_min")) * 60
        if due_backup:
            self._next_backup = now + clamp_interval(cfg.get("backup_interval_min")) * 60
            self._force_backup = False
        self._status_at = now

        self._update(busy=True)
        try:
            res = self._call(lambda: _work(cfg, due_save, due_backup), timeout=40.0)
        except TimeoutError:
            self._update(busy=False)
            self.log("warn", "PowerPoint/WPS 无响应（可能弹着对话框），本次跳过")
            return
        except Exception as exc:
            self._update(busy=False, attached=False, docs=0, doc_names=[],
                         error=str(exc))
            self.log("error", "连接失败：%s" % exc)
            return

        self._apply(res, due_save, due_backup)

    def _apply(self, res, due_save, due_backup):
        now = time.time()
        docs = res.get("docs") or []
        app_name = res.get("app") or ""
        errors = res.get("errors") or []
        notes = res.get("notes") or []
        saved = res.get("saved") or []
        backups = res.get("backups") or []

        if app_name:
            self._update(attached=True, app=app_name, docs=len(docs),
                         doc_names=[d["name"] for d in docs],
                         error=(errors[0] if errors else ""))
        else:
            self._update(attached=False, docs=0, doc_names=[],
                         error=(errors[0] if errors else ""))

        if saved:
            self._update(last_save=now)
            self.log("ok", "已自动保存：%s" % "、".join(saved))
        if backups:
            self._update(last_backup=now)
            extra = ""
            if res.get("removed"):
                extra = "（清理旧备份 %d 份）" % res["removed"]
            self.log("ok", "已备份：%s%s" % ("、".join(backups), extra))

        # 只在“什么都没做”的时候说一句，避免刷屏
        if due_save and not saved and not due_backup:
            if notes:
                self.log("note", notes[0])
            elif all(d.get("saved") for d in docs) and docs:
                self.log("note", "文档没有改动，未执行保存")
        if due_backup and not backups and not errors:
            if notes:
                self.log("note", notes[0])
        for err in errors[1:]:
            self.log("error", err)

        sig = (app_name, len(docs), tuple(d["name"] for d in docs))
        if sig != self._last_sig:
            self._last_sig = sig
            if app_name and docs:
                self.log("info", "已连接 %s，当前 %d 个文档" % (app_name, len(docs)))
            elif app_name:
                self.log("info", "已连接 %s，暂无打开的文档" % app_name)
        self._update(busy=False)

    def _call(self, fn, timeout):
        box = {}

        def run():
            try:
                box["v"] = fn()
            except Exception as exc:
                box["e"] = exc

        t = threading.Thread(target=run, daemon=True)
        t.start()
        t.join(timeout)
        if t.is_alive():
            raise TimeoutError("timeout")
        if "e" in box:
            raise box["e"]
        return box.get("v")


# --------------------------------------------------------------------------- #
# 独立自测：python autosave.py [--backup]
# --------------------------------------------------------------------------- #
if __name__ == "__main__":
    import sys

    cfg = load_config()
    if "--backup" in sys.argv:
        res = _work(cfg, False, True)
        print(json.dumps(res, ensure_ascii=False, indent=2))
    else:
        print("attach:", _attach()[1])
