# -*- coding: utf-8 -*-
"""
字体下载与安装模块
- 从 GitHub release 下载开源字体
- 解压 zip / 直接下载 ttf/otf
- 安装字体到 Windows（复制到 Fonts 目录 + 注册表注册）
"""
import os
import re
import json
import zipfile
import shutil
import urllib.request

from font_db import OPEN_SOURCE_FONTS
import paths


# 字体扩展名
FONT_EXTS = (".ttf", ".otf", ".ttc")


def _github_api(url):
    """带 UA 的 GitHub API 请求"""
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
        "Accept": "application/vnd.github+json",
    })
    return json.load(urllib.request.urlopen(req, timeout=20))


def _resolve_download_info(font_key: str) -> dict:
    """
    解析字体下载信息，返回：
    {
        'download_url': 直接下载的字体文件 URL（单个 ttf/otf/ttc），
        'zip_url': zip 包 URL（若无可返回 None），
        'font_names': 预期字体文件名列表（用于从 zip 里筛选）
    }
    """
    info = OPEN_SOURCE_FONTS.get(font_key)
    if not info:
        raise ValueError(f"未找到字体信息：{font_key}")

    # 从 url 提取 GitHub 仓库
    m = re.search(r"github\.com/([^/]+)/([^/]+)", info["url"])
    if not m:
        # 非 GitHub 源（如小米、阿里官网），无法自动下载
        raise ValueError(f"该字体需手动到官网下载：{info['url']}")

    owner, repo = m.group(1), m.group(2)
    api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"

    data = _github_api(api_url)
    assets = data.get("assets", [])

    # 1. 优先找单个字体文件（ttf/otf/ttc）
    font_files = [a for a in assets
                  if a["name"].lower().endswith(FONT_EXTS)]
    # 2. 找 zip 包
    zip_files = [a for a in assets if a["name"].lower().endswith(".zip")]

    result = {"download_url": None, "zip_url": None, "font_names": []}

    if font_files:
        # 优先选 Regular/非 Mono/非 Bold 的常规字重
        preferred = None
        for f in font_files:
            n = f["name"]
            if "regular" in n.lower() or "medium" in n.lower():
                preferred = f
                break
        if preferred is None:
            preferred = font_files[0]
        # 用 asset 的 API URL（api.github.com 走代理更稳定）
        result["download_url"] = preferred.get("url") or preferred["browser_download_url"]
        result["font_names"] = [preferred["name"]]

    if zip_files:
        # 优先选简体中文（SC/CN）且非等宽（HW/Mono）的包
        def zip_score(a):
            n = a["name"].lower()
            score = 0
            if "hw" in n or "mono" in n:
                score += 1000  # 等宽版本靠后（非常规排版）
            if "sc" in n or "cn" in n or "zh" in n:
                score -= 1000  # 简体中文优先
            if "otc" in n:
                score -= 100  # TTC 超级集合优先
            score += a.get("size", 0) // (1024 * 1024)  # 越大约靠后
            return score
        zip_files.sort(key=zip_score)
        result["zip_url"] = zip_files[0].get("url") or zip_files[0]["browser_download_url"]

    if not result["download_url"] and not result["zip_url"]:
        raise ValueError(f"未找到可下载的字体文件，请到官网手动下载：{info['url']}")

    return result


def _download_file(url, dest_path, progress_cb=None):
    """下载文件，带进度回调，手动处理重定向保留 header"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/120.0 Safari/537.36",
        "Accept": "*/*",
        "Accept-Encoding": "identity",
    }
    # GitHub asset API 需要 octet-stream 才返回文件内容
    if "api.github.com" in url and "/assets/" in url:
        headers["Accept"] = "application/octet-stream"

    current_url = url
    for _ in range(5):  # 最多跟随 5 次重定向
        req = urllib.request.Request(current_url, headers=headers)
        resp = urllib.request.urlopen(req, timeout=60)
        # 检查是否重定向
        final = resp.geturl()
        if final == current_url:
            break
        # 跟随重定向
        resp.close()
        current_url = final

    total = int(resp.headers.get("Content-Length", 0))
    downloaded = 0
    with paths.open_binary(dest_path, "wb") as f:
        while True:
            chunk = resp.read(65536)
            if not chunk:
                break
            f.write(chunk)
            downloaded += len(chunk)
            if progress_cb and total > 0:
                progress_cb(downloaded, total)
    resp.close()
    return dest_path


def _find_font_files_in_dir(d) -> list:
    """递归查找目录下所有字体文件"""
    result = []
    for root, dirs, files in os.walk(d):
        for fn in files:
            if fn.lower().endswith(FONT_EXTS):
                result.append(os.path.join(root, fn))
    return result


def download_font(font_key: str, dest_dir: str, progress_cb=None) -> list:
    """
    下载字体，返回下载的字体文件路径列表。
    progress_cb(msg: str) 阶段消息回调
    """
    info = _resolve_download_info(font_key)

    # 下载目录必须真实可写：用户选的目录可能已被删除、在只读盘上、或超长
    dest_dir = paths.normalize(dest_dir)
    if not dest_dir or not paths.ensure_dir(dest_dir) or not paths.is_writable(dest_dir):
        dest_dir = paths.temp_dir("font_dl_")
    downloaded_files = []

    # 下载字节进度回调（可选，节流：每 256KB 回调一次）
    def _make_dl_cb(label):
        last = [0]

        def cb(cur, total):
            if progress_cb and total > 0:
                # 节流：每 256KB 或每 5% 回调一次
                if cur - last[0] >= 256 * 1024 or cur >= total:
                    last[0] = cur
                    progress_cb(f"{label} {cur*100//total}% ({cur//1024}/{total//1024} KB)")
        return cb

    # 1. 优先直接下载单个字体文件
    if info["download_url"]:
        fn = info["font_names"][0]
        dest = os.path.join(dest_dir, fn)
        _download_file(info["download_url"], dest, _make_dl_cb("下载中"))
        downloaded_files.append(dest)
        return downloaded_files

    # 2. 下载 zip 包并解压
    if info["zip_url"]:
        zip_path = os.path.join(dest_dir, "font_package.zip")
        _download_file(info["zip_url"], zip_path, _make_dl_cb("下载压缩包"))

        if progress_cb:
            progress_cb("解压中…")
        extract_dir = os.path.join(dest_dir, "_extract")
        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(extract_dir)

        # 筛选字体文件（优先 ttf/otf，排除 Mono 变体）
        font_files = _find_font_files_in_dir(extract_dir)
        if not font_files:
            raise ValueError("压缩包中未找到字体文件")

        # 优先选非 Mono、非 Bold 的常规字重
        font_files.sort(key=lambda p: (
            "mono" in os.path.basename(p).lower(),
            "bold" in os.path.basename(p).lower(),
            os.path.getsize(p),
        ))

        # 复制到 dest_dir（取前几个常规字重，避免重复）
        # 简化：复制所有非 Mono 的字体，或最多 3 个
        copied = []
        for fp in font_files:
            base = os.path.basename(fp)
            if "mono" in base.lower():
                continue
            dest = os.path.join(dest_dir, base)
            shutil.copy2(fp, dest)
            copied.append(dest)
            if len(copied) >= 3:
                break
        if not copied:
            # 全是 Mono，就复制第一个
            base = os.path.basename(font_files[0])
            dest = os.path.join(dest_dir, base)
            shutil.copy2(font_files[0], dest)
            copied.append(dest)

        # 清理
        shutil.rmtree(paths.long_path(extract_dir), ignore_errors=True)
        try:
            os.remove(paths.long_path(zip_path))
        except Exception:
            pass

        downloaded_files.extend(copied)
        return downloaded_files

    raise ValueError("无法解析下载地址")


def install_font(font_file: str, progress_cb=None) -> bool:
    """
    安装字体到本机字体库。
    Windows：复制到用户字体目录 + 注册表（再尝试系统 Fonts）
    macOS  ：复制到 ~/Library/Fonts（只影响当前用户，不需要 sudo）
    """
    font_file = paths.resolve_existing(font_file)
    if not paths.is_file(font_file):
        reasons = "；".join(paths.describe_missing(font_file))
        raise FileNotFoundError(f"字体文件不存在：{font_file}\n{reasons}")

    ext = os.path.splitext(font_file)[1].lower()
    if ext not in FONT_EXTS:
        raise ValueError(f"不是字体文件：{font_file}")

    font_name = os.path.basename(font_file)

    if paths.MACOS:
        target_dir = paths.user_fonts_dir()
        if not target_dir or not paths.ensure_dir(target_dir):
            raise RuntimeError(
                "无法创建字体目录：\n~/Library/Fonts\n\n"
                "常见原因：磁盘只读、权限被策略限制。")
        target = os.path.join(target_dir, font_name)
        shutil.copy2(paths.long_path(font_file), paths.long_path(target))
        # 让正在运行的字体服务立刻看到新字体（失败不影响安装本身）
        try:
            import subprocess
            subprocess.run(["/usr/bin/atsutil", "fonts", "-update"],
                           stdout=subprocess.DEVNULL,
                           stderr=subprocess.DEVNULL, timeout=8)
        except Exception:
            pass
        if progress_cb:
            progress_cb(f"已安装：{font_name} → ~/Library/Fonts")
        return True

    reg_path = r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    value_name = _font_reg_name(font_name)

    # 方案1：复制到用户字体目录（无需管理员）
    user_fonts_dir = paths.user_fonts_dir()
    if not paths.ensure_dir(user_fonts_dir) or not paths.is_writable(user_fonts_dir):
        raise RuntimeError(
            "无法写入用户字体目录：\n%s\n\n"
            "常见原因：目录被安全软件锁定、磁盘只读、或用户配置目录异常。" % user_fonts_dir)
    target = os.path.join(user_fonts_dir, font_name)
    shutil.copy2(paths.long_path(font_file), paths.long_path(target))

    # 注册表注册（HKCU）
    try:
        import winreg
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path) as key:
            winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ, target)
    except Exception:
        pass

    # 方案2：尝试安装到系统 Fonts 目录（需管理员，失败则忽略，用户字体目录已够用）
    try:
        win_dir = paths.known_folder(paths.FOLDERID_WINDOWS,
                                     os.environ.get("WINDIR", "C:\\Windows"))
        sys_fonts = os.path.join(win_dir, "Fonts")
        if paths.is_dir(sys_fonts):
            shutil.copy2(paths.long_path(font_file),
                         paths.long_path(os.path.join(sys_fonts, font_name)))
            import winreg
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                winreg.SetValueEx(key, value_name, 0, winreg.REG_SZ,
                                  os.path.join(sys_fonts, font_name))
    except Exception:
        pass

    # 通知系统刷新字体
    try:
        import ctypes
        ctypes.windll.user32.SendMessageTimeoutW(
            0xFFFF, 0x001D, 0, 0, 0x0002, 5000, None)  # WM_FONTCHANGE
    except Exception:
        pass

    if progress_cb:
        progress_cb(f"已安装：{font_name}")
    return True


def _font_reg_name(filename: str) -> str:
    """生成注册表字体名（去掉扩展名，加 "(TrueType)" 后缀）"""
    base = os.path.splitext(filename)[0]
    ext = os.path.splitext(filename)[1].lower()
    suffix = "(OpenType)" if ext == ".otf" else "(TrueType)"
    return f"{base} {suffix}"


def download_and_install(font_key: str, progress_cb=None) -> dict:
    """
    完整流程：下载 + 安装。
    progress_cb(msg: str) 阶段消息回调
    """
    tmp_dir = paths.temp_dir("font_dl_")
    try:
        if progress_cb:
            progress_cb("解析下载地址…")

        files = download_font(font_key, tmp_dir, progress_cb=progress_cb)

        installed = []
        for fp in files:
            if progress_cb:
                progress_cb(f"安装中：{os.path.basename(fp)}…")
            install_font(fp)
            installed.append(os.path.basename(fp))

        return {"ok": True, "installed": installed, "count": len(installed)}
    finally:
        shutil.rmtree(paths.long_path(tmp_dir), ignore_errors=True)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        key = sys.argv[1]
        r = download_and_install(key, progress_cb=lambda m: print(m))
        print("\n结果:", r)
