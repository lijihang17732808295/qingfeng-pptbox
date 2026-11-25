# -*- coding: utf-8 -*-
"""
功能自检：用内置的 3 份测试 PPT 验证整条链路。

为什么需要它：用户报“路径不存在”时，问题可能出在资源释放、临时目录、
字体目录、Office 定位、长路径、COM 调用中的任何一环。让用户自己找一份
PPT 来复现，往往连输入都拿不到。内置样本保证自检永远有确定的输入，
每一步都有明确期望值，结论可以直接看出是哪一环坏了。

一次自检依次验证：
  1 资源释放   内置 PPT 能否被正确解压落盘（换机器/杀软拦截最容易坏在这）
  2 路径层     程序目录、临时目录、用户数据目录、长路径、中文空格路径
  3 字体索引   能否枚举本机字体（缓存是否命中）
  4 字体检测   对 3 份样本做真实解析，并核对预期字体是否都在
  5 自动保存   能否附着 PowerPoint / WPS，配置能否读写
  6 图片化导出 真实导出一次（需装 Office 或 WPS）

结论最后写入用户数据目录，便于把报告发给开发者。
"""

import base64
import gzip
import json
import os
import shutil
import time
import traceback
import zipfile

import paths
import selftest_data

ST_OK = "ok"
ST_WARN = "warn"
ST_FAIL = "fail"
ST_SKIP = "skip"

STATUS_TEXT = {ST_OK: "通过", ST_WARN: "注意", ST_FAIL: "失败", ST_SKIP: "跳过"}


def _step(key, title, status, lines=None):
    return {"key": key, "title": title, "status": status,
            "lines": [str(x) for x in (lines or [])]}


def _secs(t0):
    d = time.time() - t0
    return "%.1f 秒" % d if d >= 1 else "%d 毫秒" % max(1, int(d * 1000))


def _short(exc, limit=160):
    text = str(exc).replace("\r", " ").replace("\n", " ").strip()
    if len(text) > limit:
        text = text[:limit] + "…"
    return "%s: %s" % (exc.__class__.__name__, text) if text else exc.__class__.__name__


def _purge_old_work(days=1):
    """清掉上一次自检留下的临时目录，避免长期使用后 Temp 里堆一串样本。"""
    try:
        import glob
        import shutil
        import tempfile
        root = tempfile.gettempdir()
        deadline = time.time() - days * 86400
        for d in glob.glob(os.path.join(root, "selftest_*")):
            try:
                if os.path.isdir(paths.long_path(d)) and os.path.getmtime(paths.long_path(d)) < deadline:
                    shutil.rmtree(paths.long_path(d), ignore_errors=True)
            except Exception:
                pass
    except Exception:
        pass


# ---------------------------------------------------------------- 资源释放
def extract_samples(work_dir=None, emit=None):
    """把内置 PPT 解压到 work_dir，返回 (工作目录, [(meta, 路径), ...])。"""
    _purge_old_work()
    work = work_dir or paths.temp_dir("selftest_")
    paths.ensure_dir(work)
    out = []
    for item in selftest_data.SELFTEST_FILES:
        if emit:
            emit("phase", "释放测试文件：%s" % item["name"])
        raw = gzip.decompress(base64.b64decode(item["b64"]))
        if len(raw) != item["size"]:
            raise RuntimeError("%s 解压后大小不符（%d ≠ %d）"
                               % (item["name"], len(raw), item["size"]))
        target = os.path.join(work, item["name"])
        with paths.open_binary(target, "wb") as f:
            f.write(raw)
        out.append((item, target))
    return work, out


def _step_bundle(steps, emit):
    emit("phase", "① 正在释放内置测试文件…")
    t0 = time.time()
    try:
        work, pairs = extract_samples(emit=emit)
    except Exception as exc:
        steps.append(_step("bundle", "资源释放", ST_FAIL,
                           ["无法解压内置测试文件：%s" % _short(exc)]))
        return []
    lines, bad = [], 0
    for item, target in pairs:
        try:
            ok_size = os.path.getsize(paths.long_path(target)) == item["size"]
            with zipfile.ZipFile(paths.long_path(target)) as z:
                names = z.namelist()
                slides = len([n for n in names if n.startswith("ppt/slides/slide")
                              and n.endswith(".xml")])
                has_main = "ppt/presentation.xml" in names
            good = ok_size and has_main and slides == item["slides"]
            if not good:
                bad += 1
                lines.append("%s：大小/页数不符（%d 页，期望 %d 页）"
                             % (item["name"], slides, item["slides"]))
            else:
                lines.append("%s：%d 页 · %.1f KB"
                             % (item["name"], slides, item["size"] / 1024.0))
        except Exception as exc:
            bad += 1
            lines.append("%s：读取失败 %s" % (item["name"], _short(exc, 90)))
    lines.append("释放位置：%s（%s）" % (work, _secs(t0)))
    if bad:
        steps.append(_step("bundle", "资源释放", ST_FAIL, lines))
    else:
        steps.append(_step("bundle", "资源释放", ST_OK, lines))
    return pairs


# ---------------------------------------------------------------- 路径层
def _step_paths(steps, emit):
    emit("phase", "② 正在检查路径层…")
    t0 = time.time()
    rows = []
    probe_dirs = []

    def add(name, ok, detail="", soft=False):
        rows.append((name, bool(ok), str(detail), soft))

    add("程序目录", paths.is_dir(paths.exe_dir()), paths.exe_dir())
    try:
        td = paths.temp_dir("selftest_t_")
        probe_dirs.append(td)
        add("临时目录可写", paths.is_writable(td), td)
    except Exception as exc:
        add("临时目录可写", False, _short(exc))
    ud = paths.user_data_dir()
    add("用户数据目录可写", paths.is_writable(ud), ud)
    doc = paths.known_folder(getattr(paths, "FOLDERID_DOCUMENTS", ""), "")
    add("系统文档目录", bool(doc), doc or "未能解析（备份目录会退回用户目录）", soft=True)

    # 长路径：> 260 字符仍能写入/读回/删除
    try:
        base = paths.temp_dir("selftest_long_")
        probe_dirs.append(base)
        deep = base
        for i in range(12):
            deep = os.path.join(deep, "seg_%02d_abcdefghijklmnop" % i)
        paths.ensure_dir(deep)
        probe = os.path.join(deep, "长路径探针.pptx")
        payload = b"selftest-" + b"x" * 64
        with paths.open_binary(probe, "wb") as f:
            f.write(payload)
        with paths.open_binary(probe, "rb") as f:
            back = f.read()
        ok = back == payload
        add("长路径（%d 字符）" % len(probe), ok,
            "读写正常" if ok else "读回内容不一致")
        try:
            with paths.open_binary(probe, "rb"):
                pass
            os.remove(paths.long_path(probe))
        except Exception:
            pass
    except Exception as exc:
        add("长路径（>260 字符）", False, _short(exc))

    # 中文 + 空格 + 引号 的相对/绝对路径还原
    try:
        base = paths.temp_dir("selftest_cn_")
        probe_dirs.append(base)
        cn = os.path.join(base, "中文 目录", "测试 文件.pptx")
        paths.ensure_dir(os.path.dirname(cn))
        with paths.open_binary(cn, "wb") as f:
            f.write(b"x")
        got = paths.resolve_existing('  "%s"  ' % cn)
        add("中文/空格/带引号路径", paths.same_file(got, cn),
            got if paths.same_file(got, cn) else "还原为 %s" % got)
    except Exception as exc:
        add("中文/空格/带引号路径", False, _short(exc))

    for d in probe_dirs:
        try:
            shutil.rmtree(paths.long_path(d), ignore_errors=True)
        except Exception:
            pass

    ok_all = all(r[1] for r in rows if not r[3])
    soft_bad = [r for r in rows if r[3] and not r[1]]
    lines = ["%s %s — %s" % ("✓" if ok else ("!" if soft else "✗"), name, detail)
             for name, ok, detail, soft in rows]
    lines.append("耗时 %s" % _secs(t0))
    status = ST_OK
    if not ok_all:
        status = ST_FAIL
    elif soft_bad:
        status = ST_WARN
    steps.append(_step("paths", "路径层", status, lines))


# ---------------------------------------------------------------- 字体索引
def _step_index(steps, emit):
    emit("phase", "③ 正在加载本机字体索引…")
    t0 = time.time()
    try:
        import font_detect
        idx = font_detect.get_installed_fonts()
        count = len(idx)
        cached = bool(getattr(idx, "loaded_from_cache", False))
        detail = ["本机字体：%d 款（%s）" % (count, "命中缓存" if cached else "本次新建索引"),
                  "索引缓存：%s" % font_detect._cache_path(),
                  "耗时 %s" % _secs(t0)]
        if count <= 0:
            steps.append(_step("index", "字体索引", ST_FAIL,
                               ["没有枚举到任何字体，字体检测会无法判断“缺失”"] + detail))
        else:
            steps.append(_step("index", "字体索引", ST_OK, detail))
    except Exception as exc:
        steps.append(_step("index", "字体索引", ST_FAIL,
                           ["枚举失败：%s" % _short(exc)]))
    return steps


# ---------------------------------------------------------------- 字体检测
def _step_detect(steps, emit, pairs):
    emit("phase", "④ 正在用内置 PPT 做字体检测…")
    if not pairs:
        steps.append(_step("detect", "字体检测", ST_SKIP, ["没有可用的测试文件"]))
        return
    try:
        import font_detect
    except Exception as exc:
        steps.append(_step("detect", "字体检测", ST_FAIL,
                           ["模块导入失败：%s" % _short(exc)]))
        return
    lines, bad = [], 0
    for item, path in pairs:
        t0 = time.time()
        try:
            fonts, _src, notes = font_detect.extract_fonts(path)
            lower = {f.lower() for f in fonts}
            missing = [f for f in item["expect"] if f.lower() not in lower]
            if missing:
                bad += 1
                lines.append("%s：缺 %s（共 %d 款，%s）"
                             % (item["name"], "、".join(missing), len(fonts), _secs(t0)))
            else:
                lines.append("%s：%d 款字体，预期 %d 款全部命中（%s）"
                             % (item["name"], len(fonts), len(item["expect"]), _secs(t0)))
            if notes:
                lines.append("   备注：%s" % "；".join(str(n) for n in notes)[:100])
        except Exception as exc:
            bad += 1
            lines.append("%s：解析失败 %s" % (item["name"], _short(exc, 110)))
    steps.append(_step("detect", "字体检测", ST_FAIL if bad else ST_OK, lines))


# ---------------------------------------------------------------- 自动保存
def _step_autosave(steps, emit):
    emit("phase", "⑤ 正在检查自动保存引擎…")
    try:
        import autosave
    except Exception as exc:
        steps.append(_step("autosave", "自动保存", ST_FAIL,
                           ["模块导入失败：%s" % _short(exc)]))
        return

    lines = []
    status = ST_OK
    try:
        cfg = autosave.load_config()
        need = ("save_enabled", "backup_enabled", "save_interval_min",
                "backup_interval_min", "backup_dir", "keep_copies")
        missing = [k for k in need if k not in cfg]
        lines.append("配置文件：%s%s" % (
            autosave.config_path(), "" if not missing else "（缺键 %s）" % ",".join(missing)))
        lines.append("备份目录：%s" % cfg.get("backup_dir", ""))
        if missing:
            status = ST_WARN
    except Exception as exc:
        status = ST_WARN
        lines.append("配置读取异常：%s" % _short(exc))

    try:
        apps = autosave._apps()
        docs = autosave._all_docs(apps)
        labels = "、".join(label for label, _ in apps)
        lines.append("已附着：%s（%d 个文档）" % (labels, len(docs)))
        for d in docs[:4]:
            lines.append("   · %s  [%s]%s" % (
                d.get("name", "?"), d.get("app", "?"),
                "  只读" if d.get("readonly") else ""))
        if len(docs) > 4:
            lines.append("   · 其余 %d 个已省略" % (len(docs) - 4))
    except Exception as exc:
        lines.append("未附着到 PowerPoint / WPS：%s" % _short(exc, 90))
        lines.append("不影响字体检测与图片化；需要自动保存时再打开演示程序即可。")
        if status == ST_OK:
            status = ST_SKIP
    steps.append(_step("autosave", "自动保存", status, lines))


# ---------------------------------------------------------------- 图片化
def _step_flatten(steps, emit, pairs, work):
    emit("phase", "⑥ 正在真实导出一次页面图片化（10–30 秒）…")
    if not pairs:
        steps.append(_step("flatten", "图片化导出", ST_SKIP, ["没有可用的测试文件"]))
        return
    try:
        import slide_export
    except Exception as exc:
        steps.append(_step("flatten", "图片化导出", ST_FAIL,
                           ["模块导入失败：%s" % _short(exc)]))
        return

    engine = ""
    if paths.MACOS:
        try:
            if slide_export.find_libreoffice():
                engine = "LibreOffice"
            elif slide_export.find_powerpoint_exe():
                engine = "Microsoft PowerPoint"
        except Exception:
            engine = ""
    else:
        try:
            exe = slide_export.find_powerpoint_exe()
            if exe:
                engine = "Microsoft PowerPoint"
        except Exception:
            exe = None
    if not engine and not paths.MACOS:
        try:
            import autosave
            labels = [l for l, _ in autosave._apps()]
            if labels:
                engine = labels[0]
        except Exception:
            pass
    if not engine:
        miss = ("本机没有可用的导出引擎",
                "图片化需要 LibreOffice 或 Microsoft PowerPoint；字体检测不受影响")
        if not paths.MACOS:
            miss = ("本机没有可用的 PowerPoint / WPS 演示",
                    "图片化功能需要它们之一；字体检测不受影响")
        steps.append(_step("flatten", "图片化导出", ST_SKIP, list(miss)))
        return

    item, src = pairs[0]
    out_dir = os.path.join(work, "导出结果")
    paths.ensure_dir(out_dir)
    out = os.path.join(out_dir, os.path.splitext(item["name"])[0] + "_图片化.pptx")
    t0 = time.time()
    try:
        slide_export.flatten_ppt(src, out, dpi=96,
                                 progress_cb=lambda m: emit("phase", "⑥ " + str(m)))
        size = os.path.getsize(paths.long_path(out))
        with zipfile.ZipFile(paths.long_path(out)) as z:
            names = z.namelist()
            slides = len([n for n in names if n.startswith("ppt/slides/slide")
                          and n.endswith(".xml")])
            media = len([n for n in names if n.startswith("ppt/media/")])
        lines = ["引擎：%s" % engine,
                 "源文件：%s（%d 页）" % (item["name"], item["slides"]),
                 "输出：%s" % out,
                 "结果：%d 页 / %d 张图片 / %.0f KB / 耗时 %s"
                 % (slides, media, size / 1024.0, _secs(t0))]
        if slides == item["slides"] and media >= item["slides"] and size > 0:
            steps.append(_step("flatten", "图片化导出", ST_OK, lines))
        else:
            steps.append(_step("flatten", "图片化导出", ST_FAIL,
                               lines + ["页数或内嵌图片数量不符，导出结果可能不完整"]))
        try:
            os.remove(paths.long_path(out))
        except Exception:
            pass
    except Exception as exc:
        steps.append(_step("flatten", "图片化导出", ST_FAIL,
                           ["引擎：%s" % engine,
                            "导出失败：%s" % _short(exc, 220),
                            "若提示找不到文件，多为文件被占用或输出目录不可写"]))


# ---------------------------------------------------------------- 报告
def _system_line():
    """系统描述（Windows / macOS 通用）"""
    try:
        import platform
        bits = platform.machine()
        return "%s %s %s" % (platform.system(), platform.release(), bits)
    except Exception:
        return os.environ.get("OS", "未知")


def _report_text(steps, work, elapsed):
    import app_info
    lines = [
        "PPT 工具箱 · 功能自检报告",
        "=" * 46,
        "时间：%s" % time.strftime("%Y-%m-%d %H:%M:%S"),
        "版本：v%s" % app_info.VERSION,
        "模式：%s" % (("打包 APP" if paths.MACOS else "打包 EXE")
                      if paths.is_frozen() else "源码运行"),
        "程序：%s" % (paths.exe_path() or "-"),
        "系统：%s" % _system_line(),
        "Python：%s" % __import__("sys").version.split()[0],
        "测试文件目录：%s" % work,
        "耗时：%s" % elapsed,
        "",
    ]
    for i, s in enumerate(steps, 1):
        lines.append("[%s] %d. %s" % (STATUS_TEXT[s["status"]], i, s["title"]))
        for ln in s["lines"]:
            lines.append("      " + ln)
        lines.append("")
    ok = sum(1 for s in steps if s["status"] == ST_OK)
    warn = sum(1 for s in steps if s["status"] == ST_WARN)
    fail = sum(1 for s in steps if s["status"] == ST_FAIL)
    skip = sum(1 for s in steps if s["status"] == ST_SKIP)
    lines.append("=" * 46)
    lines.append("结论：%d 项通过，%d 项失败，%d 项注意，%d 项跳过" % (ok, fail, warn, skip))
    if fail:
        lines.append("存在失败项，请把本报告与上面的失败步骤一并发给开发者。")
    elif warn or skip:
        lines.append("主要功能可用；注意项见上，通常与环境配置有关。")
    else:
        lines.append("全部检查通过，功能链路正常。")
    return "\n".join(lines)


def run(cb=None, work_dir=None):
    """执行自检。cb(kind, payload)，kind ∈ phase / step / done。"""
    def emit(kind, payload):
        if cb:
            try:
                cb(kind, payload)
            except Exception:
                pass

    started = time.time()
    steps = []
    work = ""

    emit("phase", "准备环境…")
    pythoncom = None
    try:
        import pythoncom
        pythoncom.CoInitialize()
    except Exception:
        pythoncom = None
    try:
        pairs = _step_bundle(steps, emit)
        if pairs:
            work = os.path.dirname(pairs[0][1])
        for fn in (lambda: _step_paths(steps, emit),
                   lambda: _step_index(steps, emit),
                   lambda: _step_detect(steps, emit, pairs),
                   lambda: _step_autosave(steps, emit),
                   lambda: _step_flatten(steps, emit, pairs, work or paths.temp_dir("selftest_"))):
            try:
                fn()
            except Exception:
                steps.append(_step("unknown", "未预期的检查项", ST_FAIL,
                                   [traceback.format_exc().strip().splitlines()[-1]]))
            if steps:
                emit("step", steps[-1])

        elapsed = _secs(started)
        text = _report_text(steps, work or "-", elapsed)
        target = ""
        emit("phase", "正在写入报告…")
        try:
            name = "selftest_%s.txt" % time.strftime("%Y%m%d_%H%M%S")
            target = os.path.join(paths.user_data_dir(), name)
            with paths.open_binary(target, "w", encoding="utf-8") as f:
                f.write(text)
        except Exception:
            target = ""
        report = {"steps": steps, "text": text, "work_dir": work,
                  "report_path": target, "elapsed": elapsed,
                  "fail": sum(1 for s in steps if s["status"] == ST_FAIL)}
        emit("done", report)
        return report
    finally:
        try:
            if pythoncom is not None:
                pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    def _cli(kind, payload):
        if kind == "phase":
            print("·", payload)
        elif kind == "step":
            print("[%s] %s" % (STATUS_TEXT[payload["status"]], payload["title"]))
            for ln in payload["lines"]:
                print("     ", ln)
        elif kind == "done":
            print()
            print(payload["text"])
            print("报告：", payload["report_path"] or "（写入失败）")

    run(_cli)
