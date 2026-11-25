# -*- coding: utf-8 -*-
"""
PPT 工具箱（高清抗锯齿版）
功能 1：字体检测 —— 多路解析 PPT 用到的全部字体，缺失的开源字体可一键下载安装
功能 2：页面图片化 —— 每页导出高清图片再插回新 PPT，版式彻底固定
功能 3：自动保存 —— 按设定间隔自动保存正在编辑的 PPT，并另存副本到指定目录
        （同时兼容 Microsoft PowerPoint 与 WPS 演示），可设开机自启动

UI：DPI 感知 + Pillow 超采样圆角卡片（无毛边）+ 自绘胶囊按钮
"""
import os
import re
import sys
import time
import queue
import threading
import traceback
import ctypes
import tkinter as tk
import tkinter.font as tkfont
from collections import OrderedDict
from tkinter import ttk, filedialog, messagebox, simpledialog

from PIL import Image, ImageDraw, ImageTk

import app_info
import autosave
import autostart
import font_detect
import font_install
import slide_export
import paths


# ============================================================
# DPI：必须在创建 Tk 之前调用，否则 Windows 会把整个窗口位图拉伸 → 全屏毛边
# ============================================================
def enable_dpi_awareness():
    """必须在创建 Tk 之前调用；失败则整窗被系统位图拉伸 → 全屏发虚"""
    if sys.platform != "win32":
        return
    # Win10 1703+：Per-Monitor V2 最理想
    try:
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except Exception:
        pass
    try:
        if ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0:  # PER_MONITOR_DPI_AWARE
            return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def work_area():
    """桌面可用区域（不含任务栏）。返回 (x, y, w, h)，取不到时 w 为 0。"""
    if sys.platform == "win32":
        try:
            import ctypes.wintypes as wt
            rect = wt.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(
                    0x0030, 0, ctypes.byref(rect), 0):   # SPI_GETWORKAREA
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w > 200 and h > 200:
                    return rect.left, rect.top, w, h
        except Exception:
            pass
    return 0, 0, 0, 0


def system_dpi():
    """当前系统 DPI（进程若未拿到 DPI 感知，Tk 会误报 96）"""
    if sys.platform == "win32":
        try:
            dpi = ctypes.windll.user32.GetDpiForSystem()
            if 72 <= dpi <= 480:
                return dpi
        except Exception:
            pass
        try:
            dc = ctypes.windll.user32.GetDC(0)
            dpi = ctypes.windll.gdi32.GetDeviceCaps(dc, 88)  # LOGPIXELSX
            ctypes.windll.user32.ReleaseDC(0, dc)
            if 72 <= dpi <= 480:
                return dpi
        except Exception:
            pass
    return 96


# ============================================================
# 配色
# ============================================================
BG         = "#080b11"
CARD_TOP   = "#161d2b"
CARD_BOT   = "#111724"
CARD_SOLID = "#131a26"
CARD_HEAD  = "#1a2130"
BORDER     = "#222c3e"
BORDER_HL  = "#2f3d55"
TEXT       = "#e9eefb"
DIM        = "#8b96ab"
DIM2       = "#66718a"
ACCENT     = "#35d39b"
ACCENT_D   = "#1f9c73"
BLUE       = "#4d8dff"
BLUE_D     = "#2f6fd8"
WARN       = "#ffb02e"
FAIL       = "#ff5c7a"
OK         = "#3ddc97"

UI_FONT = "Microsoft YaHei UI"

# ---- 平台差异（Windows / macOS 共用同一套界面代码）----
WINDOWS = sys.platform == "win32"
MACOS = sys.platform == "darwin"
APP_FAMILY = "Microsoft PowerPoint 与 WPS 演示" if WINDOWS else "Microsoft PowerPoint"
SAVE_KEY = "Ctrl+S" if WINDOWS else "⌘S"
LOGIN_WORD = "Windows" if WINDOWS else "macOS"
AUTOSTART_HINT = ("写当前用户注册表，不需要管理员" if WINDOWS
                  else "写当前用户的 LaunchAgent，不需要管理员密码")
# macOS 的 Tk 收不到系统拖放事件，文案改成"选择文件"
DROP_TITLE = "拖拽 PPT 文件到这里" if WINDOWS else "选择一份 PPT 开始"
DROP_TIP = ("拖入 PPT 后自动解析版式 / 母版 / 主题 / 图表 / 智能图形 / 内嵌字体"
            if WINDOWS else
            "选中 PPT 后自动解析版式 / 母版 / 主题 / 图表 / 智能图形 / 内嵌字体")
DROP_TIP2 = ("拖入 PPT 后自动开始，无需安装 Office 也能解析" if WINDOWS
             else "载入 PPT 后自动开始，无需安装 Office 也能解析")

# 图片化导出清晰度（DPI, 显示名）
DPI_OPTIONS = ((150, "标清 150"), (192, "高清 192"), (288, "超清 288"))

# 自动保存 / 定时备份的常用间隔（分钟），其余值可点「自定义」输入
SAVE_PRESETS = (1, 5, 10, 15, 30)
BACKUP_PRESETS = (5, 15, 30, 60)

# ============================================================
# 圆角卡片渲染（Pillow 超采样 → 抗锯齿）
# ============================================================
_CARD_CACHE = OrderedDict()
_CARD_CACHE_MAX = 160


def _ss_for(w, h):
    """按面积自适应超采样：小控件高倍（更锐利），大卡片低倍（更快）"""
    area = w * h
    if area <= 200000:
        return 4
    if area <= 600000:
        return 3
    return 2


def _rgb(c):
    c = c.lstrip("#")
    return tuple(int(c[i:i + 2], 16) for i in (0, 2, 4))


def _mix(c1, c2, t):
    a, b = _rgb(c1), _rgb(c2)
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))


def _local_bg(card_top, card_bot, y, h):
    """取卡片在某一高度的底色（用于给嵌套卡片/贴片做圆角外填充）"""
    return _hex(_mix(card_top, card_bot, max(0.0, min(1.0, y / max(1.0, float(h))))))


def _hex(t):
    return "#%02x%02x%02x" % (t[0], t[1], t[2])


def _shift(c, amount):
    """向白(正)或黑(负)偏移，返回 hex（必须与 _mix 的入参类型一致）"""
    a = _rgb(c)
    if amount >= 0:
        return _hex(tuple(min(255, int(v + (255 - v) * amount)) for v in a))
    return _hex(tuple(max(0, int(v * (1 + amount))) for v in a))


def resource_path(*parts):
    """资源路径：onefile 在 _MEIPASS，onedir 在 exe 目录，源码在脚本目录。

    统一交给 paths.resource_path 做多候选探测，换机器/换打包方式都不会失联。
    """
    return paths.resource_path(*parts)


def card_image(w, h, r, top, bot=None, border=None, border_w=1):
    """生成抗锯齿圆角卡片（带竖向渐变 + 1px 描边）"""
    if w <= 1 or h <= 1:
        w, h = max(2, w), max(2, h)
    key = (w, h, r, top, bot, border, border_w)
    img = _CARD_CACHE.get(key)
    if img is not None:
        _CARD_CACHE.move_to_end(key)
        return img

    ss = _ss_for(w, h)
    W, H = w * ss, h * ss
    base = Image.new("RGB", (W, H))
    if bot and bot != top:
        # 用 1×H 的渐变条放大，避免逐行画线（大卡片快 5~10 倍）
        strip = Image.new("RGB", (1, H))
        px = strip.load()
        for y in range(H):
            px[0, y] = _mix(top, bot, y / max(1, H - 1))
        base = strip.resize((W, H), Image.NEAREST)
    else:
        ImageDraw.Draw(base).rectangle([0, 0, W, H], fill=_rgb(top))

    mask = Image.new("L", (W, H), 0)
    ImageDraw.Draw(mask).rounded_rectangle(
        [0, 0, W - 1, H - 1], radius=max(1, r * ss), fill=255)

    out = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    out.paste(base, (0, 0), mask)

    if border:
        bd = ImageDraw.Draw(out)
        bw = max(1, border_w * ss)
        bd.rounded_rectangle(
            [bw // 2, bw // 2, W - 1 - bw // 2, H - 1 - bw // 2],
            radius=max(1, r * ss), outline=_rgb(border), width=bw)

    photo = ImageTk.PhotoImage(out.resize((w, h), Image.LANCZOS if ss >= 3 else Image.BILINEAR))
    _CARD_CACHE[key] = photo
    while len(_CARD_CACHE) > _CARD_CACHE_MAX:
        _CARD_CACHE.popitem(last=False)
    return photo


_CORNER_CACHE = {}
_CORNER_ARCS = {"tl": (180, 270), "tr": (270, 360), "bl": (90, 180), "br": (0, 90)}
_CORNER_CENTER = {"tl": (1, 1), "tr": (0, 1), "bl": (1, 0), "br": (0, 0)}


def _corner_tile(r, c_top, c_bot, border, bg, corner):
    """r×r 圆角贴片：弧内跟随卡片渐变，弧外填背景色，边缘 4 倍超采样抗锯齿"""
    key = (r, c_top, c_bot, border, bg, corner)
    photo = _CORNER_CACHE.get(key)
    if photo is not None:
        return photo
    ss = 4
    R = max(2, int(r))
    W = R * ss
    if c_top == c_bot:
        img = Image.new("RGB", (W, W), c_top if isinstance(c_top, tuple) else _rgb(c_top))
    else:
        strip = Image.new("RGB", (1, W))
        px = strip.load()
        top_rgb, bot_rgb = _rgb(c_top), _rgb(c_bot)
        for yy in range(W):
            t = yy / max(1, W - 1)
            px[0, yy] = (int(top_rgb[0] + (bot_rgb[0] - top_rgb[0]) * t),
                         int(top_rgb[1] + (bot_rgb[1] - top_rgb[1]) * t),
                         int(top_rgb[2] + (bot_rgb[2] - top_rgb[2]) * t))
        img = strip.resize((W, W), Image.NEAREST)

    fx, fy = _CORNER_CENTER[corner]
    cx, cy = (W if fx else 0), (W if fy else 0)
    rw = W - ss // 2
    box = [cx - rw, cy - rw, cx + rw, cy + rw]

    # 弧外填「父级底色」形成圆角（Tk 画布不支持图片 alpha 合成，故预先合成）
    mask = Image.new("L", (W, W), 0)
    ImageDraw.Draw(mask).ellipse(box, fill=255)
    out = Image.composite(img, Image.new("RGB", (W, W), _rgb(bg)), mask)

    if border:
        a0, a1 = _CORNER_ARCS[corner]
        ImageDraw.Draw(out).arc(box, a0, a1, fill=_rgb(border), width=ss)

    photo = ImageTk.PhotoImage(out.resize((R, R), Image.LANCZOS))
    _CORNER_CACHE[key] = photo
    return photo


def draw_card(c, tag, x, y, w, h, r, top, bot=None, border=BORDER, bands=64, bg=None):
    """Canvas 圆角卡片：渐变横带 + 四角超采样贴片（不生成整图，缩放不卡）"""
    if w < 6 or h < 6:
        return
    bot = bot or top
    r = max(1, min(int(r), int(w // 2) - 1, int(h // 2) - 1))
    bg = bg or c.cget("bg")
    hh = float(h)

    # 1) 渐变横带（全宽，四角随后被贴片覆盖）
    n = max(1, min(bands, max(1, int(h // 2))))
    step = hh / n
    for i in range(n):
        y0 = y + round(i * step)
        y1 = y + round((i + 1) * step)
        c.create_rectangle(x, y0, x + w, y1, outline="",
                           fill=_hex(_mix(top, bot, i / max(1, n - 1))), tags=(tag,))

    # 2) 四边描边（中间直线段）
    if border:
        bx0, bx1 = x + r, x + w - r
        by0, by1 = y + r, y + h - r
        c.create_line(bx0, y + 0.5, bx1, y + 0.5, fill=border, tags=(tag,))
        c.create_line(bx0, y + h - 0.5, bx1, y + h - 0.5, fill=border, tags=(tag,))
        c.create_line(x + 0.5, by0, x + 0.5, by1, fill=border, tags=(tag,))
        c.create_line(x + w - 0.5, by0, x + w - 0.5, by1, fill=border, tags=(tag,))

    # 3) 四角贴片（弧内渐变与横带一致，不会出现拼接色块）
    top_hi = _hex(_mix(top, bot, min(1.0, r / hh)))
    bot_lo = _hex(_mix(top, bot, max(0.0, 1 - r / hh)))
    for corner, ox, oy, c1, c2 in (
            ("tl", 0, 0, top, top_hi), ("tr", 1, 0, top, top_hi),
            ("bl", 0, 1, bot_lo, bot), ("br", 1, 1, bot_lo, bot)):
        photo = _corner_tile(r, c1, c2, border, bg, corner)
        cx = x if ox == 0 else x + w - r
        cy = y if oy == 0 else y + h - r
        c.create_image(cx, cy, anchor="nw", image=photo, tags=(tag,))


# ============================================================
# 主程序
# ============================================================
class PPTToolbox:
    def __init__(self, root):
        self.root = root
        try:
            tk_dpi = float(root.winfo_fpixels("1i"))
        except Exception:
            tk_dpi = 96.0
        self.DPI = max(tk_dpi, float(system_dpi()))
        self.S = max(1.0, min(3.0, self.DPI / 96.0))

        global UI_FONT
        fams = set(tkfont.families())
        for f in ("Microsoft YaHei UI", "微软雅黑", "Microsoft YaHei",
                  "PingFang SC", "Hiragino Sans GB", "Heiti SC",
                  "Segoe UI", "SimHei"):
            if f in fams:
                UI_FONT = f
                break

        self.root.title(f"{app_info.APP_NAME} v{app_info.VERSION} · 字体检测 & 图片化 & 自动保存")
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        win_w = min(self.px(980), int(sw * 0.94))
        win_h = min(self.px(850), int(sh * 0.90))
        # 放在工作区（去掉任务栏）里居中：否则窗口底边会压在任务栏下面，
        # 底部按钮点不到
        wa_x, wa_y, wa_w, wa_h = work_area()
        if wa_w and wa_h:
            frame_extra = self.px(44)          # 标题栏 + 边框的粗略高度
            win_w = min(win_w, wa_w - self.px(16))
            win_h = min(win_h, wa_h - frame_extra - self.px(16))
            gx = wa_x + max(0, (wa_w - win_w) // 2)
            gy = wa_y + max(0, (wa_h - win_h - frame_extra) // 2)
            self.root.geometry(f"{win_w}x{win_h}+{gx}+{gy}")
        else:
            self.root.geometry(f"{win_w}x{win_h}")
        self.root.minsize(min(self.px(880), int(sw * 0.86)),
                          min(self.px(720), int(sh * 0.78)))
        self.root.configure(bg=BG)
        try:
            ph = self.logo_photo(64)
            if ph is not None:
                self.root.iconphoto(True, ph)
        except Exception:
            pass

        self.files = []
        self.detecting = False
        self.installing = False
        self._flattening = False
        self.fonts = []
        self.summary = {}
        self.current_mode = "detect"
        self._fout_dir = None
        self.dpi_var = tk.IntVar(value=192)
        self._prog = None
        self._fprog = None
        self._fstatus = "等待开始…"
        self._fhint = ""
        self._detect_seconds = 0.0
        self._prog_ids = []
        self._drop_key = None
        self._list_size = None
        self._dirty = {"detect", "flatten", "about", "autosave"}
        self._drop_hooked = False
        self._hook_tries = 0

        # ---- 自动保存 / 定时备份 ----
        self.as_cfg = autosave.load_config()
        self.as_engine = autosave.Engine(lambda: self.as_cfg)
        self._as_hits = []
        self._as_sig = None
        self._as_auto_state = autostart.state()

        self._event_queue = queue.Queue()
        self.tk_families = list(fams)

        self._init_style()
        self._build_ui()
        # 拖放必须在窗口完成映射后再挂钩：此前 TkChild 还没挂到 TkTopLevel 上，
        # 只能勾到内层窗口，真实拖拽会投递到外层而收不到
        for _d in (80, 300, 900, 1800):
            self.root.after(_d, self._setup_drag_drop)
        self.root.bind("<Map>", lambda e: self._setup_drag_drop(), add="+")
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(60, self._tick)
        self.root.after(700, self._as_refresh)
        # 延后建索引：先让窗口画出来，避免启动阶段抢 CPU 造成"卡"
        self.root.after(600, lambda: threading.Thread(
            target=self._warm_index, daemon=True).start())

    # ---------------- 工具 ----------------
    def _warm_index(self):
        """启动时后台建立本机字体索引，顶部显示字体总数（只读，不涉及 tkinter）"""
        try:
            idx = font_detect.get_installed_fonts(extra_names=self.tk_families)
            self._event_queue.put(("env", len(idx)))
        except Exception:
            pass

    def px(self, v):
        return int(round(v * self.S))

    def fmt(self, size, weight="normal"):
        return (UI_FONT, size, weight) if weight != "normal" else (UI_FONT, size)

    def logo_photo(self, pixels, mark=False):
        """加载内置 Logo（pixels 为目标像素边长，mark=True 用无字简化版），带缓存。"""
        cache = getattr(self, "_logo_cache", None)
        if cache is None:
            cache = self._logo_cache = {}
        key = (int(pixels), bool(mark))
        if key in cache:
            return cache[key]
        path = resource_path("assets", "logo_mark.png" if mark else "logo.png")
        if not os.path.isfile(path):
            return None
        px = max(16, int(pixels))
        im = Image.open(path).convert("RGBA").resize((px, px), Image.LANCZOS)
        ph = ImageTk.PhotoImage(im)
        cache[key] = ph
        return ph

    def _later(self, key, fn, delay=200):
        """合并高频 <Configure> 重绘：窗口拖动缩放时只在停下来后画一次"""
        attr = "_later_" + key
        old = getattr(self, attr, None)
        if old:
            try:
                self.root.after_cancel(old)
            except Exception:
                pass
        setattr(self, attr, self.root.after(delay, fn))

    def _init_style(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except Exception:
            pass
        rh = self.px(34)
        st.configure("Flat.Treeview",
                     background=CARD_SOLID, fieldbackground=CARD_SOLID,
                     foreground=TEXT, rowheight=rh, borderwidth=0,
                     relief="flat", font=self.fmt(10))
        st.layout("Flat.Treeview", [("Flat.Treeview.treearea", {"sticky": "nswe"})])
        st.configure("Flat.Treeview.Heading",
                     background=CARD_HEAD, foreground=DIM,
                     font=self.fmt(9, "bold"), borderwidth=0, relief="flat",
                     padding=(self.px(8), self.px(7)))
        st.map("Flat.Treeview.Heading", background=[("active", CARD_HEAD)])
        st.map("Flat.Treeview",
               background=[("selected", "#26344c")],
               foreground=[("selected", TEXT)])
        st.configure("Flat.Vertical.TScrollbar",
                     background="#2b3648", troughcolor=CARD_SOLID,
                     bordercolor=CARD_SOLID, arrowcolor=DIM2,
                     darkcolor=CARD_SOLID, lightcolor=CARD_SOLID,
                     width=self.px(10), relief="flat", arrowsize=self.px(10))
        st.map("Flat.Vertical.TScrollbar", background=[("active", "#3a475e")])

    # ---------------- 顶部 ----------------
    def _build_ui(self):
        head = tk.Frame(self.root, bg=BG)
        head.pack(fill="x", padx=self.px(34), pady=(self.px(22), 0))

        left = tk.Frame(head, bg=BG)
        left.pack(side="left")

        logo = tk.Canvas(left, width=self.px(38), height=self.px(38),
                         bg=BG, highlightthickness=0, bd=0)
        logo.pack(side="left", padx=(0, self.px(12)))
        ph = self.logo_photo(self.px(38), mark=True)
        if ph is not None:
            logo.create_image(0, 0, image=ph, anchor="nw")
        else:
            draw_card(logo, "lg", 0, 0, self.px(38), self.px(38), self.px(11),
                      ACCENT, ACCENT_D, border=None)
            logo.create_text(self.px(19), self.px(19), text="P", fill="#06251b",
                             font=self.fmt(17, "bold"))

        tbox = tk.Frame(left, bg=BG)
        tbox.pack(side="left")
        tk.Label(tbox, text="PPT 工具箱", bg=BG, fg=TEXT,
                 font=self.fmt(19, "bold")).pack(anchor="w")
        tk.Label(tbox, text=app_info.SUBTITLE, bg=BG, fg=DIM2,
                 font=self.fmt(9)).pack(anchor="w")

        # 右侧：本机字体数
        self.env_canvas = tk.Canvas(head, width=self.px(196), height=self.px(44),
                                    bg=BG, highlightthickness=0, bd=0)
        self.env_canvas.pack(side="right", pady=(self.px(2), 0))
        self._draw_env_chip()

        # ---- 功能切换 ----
        nav = tk.Frame(self.root, bg=BG)
        nav.pack(fill="x", padx=self.px(34), pady=(self.px(16), self.px(10)))

        self.nav_detect = self._pill(nav, "字体检测", lambda: self._switch_mode("detect"),
                                     kind="tab", w=126, h=38)
        self.nav_detect.pack(side="left", padx=(0, self.px(8)))
        self.nav_flatten = self._pill(nav, "页面图片化", lambda: self._switch_mode("flatten"),
                                      kind="tab", w=150, h=38)
        self.nav_flatten.pack(side="left", padx=(0, self.px(8)))
        self.nav_auto = self._pill(nav, "自动保存", lambda: self._switch_mode("autosave"),
                                   kind="tab", w=136, h=38)
        self.nav_auto.pack(side="left", padx=(0, self.px(8)))
        self.nav_about = self._pill(nav, "关于", lambda: self._switch_mode("about"),
                                    kind="tab", w=104, h=38)
        self.nav_about.pack(side="left")

        # 右侧版本角标
        self.ver_chip = tk.Canvas(nav, width=self.px(150), height=self.px(38),
                                  bg=BG, highlightthickness=0, bd=0)
        self.ver_chip.pack(side="right")
        self._draw_ver_chip()

        self.pages = tk.Frame(self.root, bg=BG)
        self.pages.pack(fill="both", expand=True, padx=0, pady=0)
        self.pages.grid_rowconfigure(0, weight=1)
        self.pages.grid_columnconfigure(0, weight=1)

        # 四个页面叠在同一格里，切换只做 tkraise，避免反复 pack/forget 触发布局计算
        self.detect_page = tk.Frame(self.pages, bg=BG)
        self.flatten_page = tk.Frame(self.pages, bg=BG)
        self.autosave_page = tk.Frame(self.pages, bg=BG)
        self.about_page = tk.Frame(self.pages, bg=BG)
        for _p in (self.detect_page, self.flatten_page, self.autosave_page,
                   self.about_page):
            _p.grid(row=0, column=0, sticky="nsew")

        self._build_detect_page()
        self._build_flatten_page()
        self._build_autosave_page()
        self._build_about_page()
        self._switch_mode("detect")

    def _draw_env_chip(self):
        c = self.env_canvas
        c.delete("all")
        w, h = self.px(196), self.px(44)
        draw_card(c, "bg", 0, 0, w, h, self.px(22), "#121826", "#0f1420", BORDER)
        c.create_oval(self.px(16), h // 2 - self.px(3), self.px(22), h // 2 + self.px(3),
                      fill=ACCENT, outline="")
        n = getattr(self, "_installed_total", 0)
        txt = f"本机字体 {n} 款" if n else "本机字体 -- 款"
        c.create_text(self.px(30), h // 2, text=txt, anchor="w",
                      fill=DIM, font=self.fmt(9))

    def _draw_ver_chip(self):
        c = self.ver_chip
        c.delete("all")
        w, h = self.px(150), self.px(38)
        txt = f"v{app_info.VERSION} · {app_info.TAGLINE}"
        # 自适应宽度
        c.create_text(w, h // 2, text=txt, anchor="e", fill=DIM2,
                      font=self.fmt(9))

    def _switch_mode(self, mode):
        self.current_mode = mode
        for btn, key in ((self.nav_detect, "detect"), (self.nav_flatten, "flatten"),
                         (self.nav_auto, "autosave"), (self.nav_about, "about")):
            btn.active = (key == mode)
            btn.repaint()
        {"detect": self.detect_page,
         "flatten": self.flatten_page,
         "autosave": self.autosave_page,
         "about": self.about_page}[mode].tkraise()
        if mode in self._dirty:
            self._dirty.discard(mode)
            self.root.after_idle(
                {"detect": self._draw_drop_card,
                 "flatten": self._draw_fdrop_card,
                 "autosave": self._draw_as_all,
                 "about": self._draw_about_card}[mode])
        elif mode == "autosave":
            # 每次进来都刷一遍：后台线程的最新状态可能已经变了
            self._as_sig = None
            self.root.after_idle(self._draw_as_all)
            self.root.after_idle(self._sync_as_buttons)

    # =========================================================
    # 胶囊按钮
    # =========================================================
    def _pill(self, parent, text, cmd, kind="primary", w=130, h=42, font_size=10):
        c = tk.Canvas(parent, width=self.px(w), height=self.px(h),
                      bg=BG, highlightthickness=0, bd=0, cursor="hand2")
        c.W, c.H, c.R = self.px(w), self.px(h), self.px(h // 2)
        c.kind, c.active = kind, False
        c.cmd = cmd
        c._enabled = True
        c._hover = False
        c._label = text
        c._font_size = font_size

        if kind == "primary":
            c.colors = (ACCENT, ACCENT_D, "#06251b")
        elif kind == "secondary":
            c.colors = (BLUE, BLUE_D, "#04152e")
        elif kind == "ghost":
            c.colors = ("#1b2333", "#161d2b", TEXT)
        else:  # tab
            c.colors = (ACCENT, ACCENT_D, "#06251b")

        c.repaint = lambda: self._paint_pill(c)
        c.set_text = lambda t: self._pill_text(c, t)
        c.set_enabled = lambda en: self._pill_enable(c, en)

        c.bind("<Enter>", lambda e: self._pill_hover(c, True))
        c.bind("<Leave>", lambda e: self._pill_hover(c, False))
        c.bind("<Button-1>", lambda e: c.cmd() if c._enabled else None)
        self._paint_pill(c)
        return c

    def _paint_pill(self, c):
        c.delete("all")
        W, H, R = c.W, c.H, c.R
        if c.kind == "tab" and not c.active:
            top, bot, fg = "#151c29", "#111724", DIM
            border = BORDER
        elif c.kind == "ghost":
            top, bot, fg = ("#232c3e", "#1a2130", TEXT) if c._hover else ("#1b2333", "#161d2b", TEXT)
            border = BORDER_HL
        else:
            top, bot, fg = c.colors
            if c._hover:
                top, bot = _shift(top, 0.12), _shift(bot, 0.12)
            border = None
        if not c._enabled:
            top, bot = "#1a2130", "#161d2b"
            fg = DIM2
            border = BORDER
        draw_card(c, "bg", 0, 0, W, H, R, top, bot, border=border)
        c.create_text(W // 2, H // 2, text=c._label, fill=fg,
                      font=self.fmt(c._font_size, "bold"), tags="txt")

    def _pill_hover(self, c, on):
        c._hover = on
        self._paint_pill(c)

    def _pill_text(self, c, t):
        c._label = t
        c.itemconfig("txt", text=t)

    def _pill_enable(self, c, en):
        c._enabled = en
        self._paint_pill(c)

    # =========================================================
    # 功能 1：字体检测页
    # =========================================================
    def _build_detect_page(self):
        page = self.detect_page

        self.drop_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                   height=self.px(108))
        self.drop_card.pack(fill="x", padx=self.px(34), pady=(0, self.px(14)))
        self.drop_card.bind("<Configure>", lambda e: self._later("drop", self._draw_drop_card))

        self.list_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0)
        self.list_card.pack(fill="both", expand=True, padx=self.px(34))
        self.list_card.bind("<Configure>", lambda e: self._later("list", self._draw_list_card))

        self.tree = ttk.Treeview(
            self.list_card, style="Flat.Treeview", show="headings",
            columns=("name", "status", "cat", "note", "src"), selectmode="browse")
        for cid, txt, wd, anchor in (
                ("name", "字体名称", 216, "w"),
                ("status", "状态", 92, "center"),
                ("cat", "类型", 96, "center"),
                ("note", "授权 / 说明", 320, "w"),
                ("src", "来源", 96, "center")):
            self.tree.heading(cid, text=txt, anchor=anchor)
            self.tree.column(cid, width=self.px(wd), minwidth=self.px(60),
                             anchor=anchor, stretch=(cid == "note"))

        self.vsb = ttk.Scrollbar(self.list_card, orient="vertical",
                                 style="Flat.Vertical.TScrollbar",
                                 command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        self._tip = None
        self.tree.bind("<Motion>", self._tree_motion)
        self.tree.bind("<Leave>", lambda e: self._hide_tip())

        self._tree_win = None
        self._vsb_win = None

        # 汇总
        self.sum_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                  height=self.px(86))
        self.sum_card.pack(fill="x", padx=self.px(34), pady=(self.px(14), 0))
        self.sum_card.bind("<Configure>", lambda e: self._later("sum", self._draw_sum_card))

        # 按钮
        btns = tk.Frame(page, bg=BG)
        btns.pack(pady=(self.px(16), self.px(18)))
        self.btn_choose = self._pill(btns, "选择文件", self.choose_file, "ghost", 118, 42)
        self.btn_choose.pack(side="left", padx=self.px(6))
        self.btn_detect = self._pill(btns, "检测字体", self.detect_fonts, "primary", 132, 42)
        self.btn_detect.pack(side="left", padx=self.px(6))
        self.btn_install = self._pill(btns, "一键装缺字", self.install_all, "secondary", 146, 42)
        self.btn_install.pack(side="left", padx=self.px(6))
        self.btn_clear = self._pill(btns, "清空", self.clear_all, "ghost", 96, 42)
        self.btn_clear.pack(side="left", padx=self.px(6))

    def _draw_drop_card(self, force=False):
        if self.current_mode != "detect":
            self._dirty.add("detect")
            return
        c = self.drop_card
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        name = getattr(self, "_loaded_name", None)
        if self.detecting and getattr(self, "_hint_text", None):
            title, sub, hint, hc = name or "正在检测…", "正在深度解析文档结构", self._hint_text, BLUE
        elif name:
            title, sub, hint, hc = name, "已载入 · 可再次检测", "检测完成后可在下方查看明细", OK
        else:
            title = DROP_TITLE
            sub = "支持 .pptx / .pptm / .ppt · 也可点下方「选择文件」"
            hint = "深度解析：版式 / 母版 / 主题 / 图表 / 智能图形 / 内嵌字体"
            hc = DIM2

        key = (w, h, title, sub, hint, hc)
        if key != getattr(self, "_drop_key", None) or force:
            self._drop_key = key
            c.delete("all")
            self._prog_ids = []
            draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)
            cx = w // 2
            c.create_text(cx, h // 2 - self.px(16), text=title, fill=TEXT,
                          font=self.fmt(14, "bold"))
            c.create_text(cx, h // 2 + self.px(8), text=sub, fill=DIM, font=self.fmt(9))
            c.create_text(cx, h // 2 + self.px(28), text=hint, fill=hc, font=self.fmt(8))
        self._draw_drop_prog()

    def _draw_drop_prog(self):
        """只动进度条两个矩形，不重建整张卡片"""
        c = self.drop_card
        w, h = c.winfo_width(), c.winfo_height()
        ids = getattr(self, "_prog_ids", [])
        if self._prog is None:
            if ids:
                for i in ids:
                    c.delete(i)
                self._prog_ids = []
            return
        pad = self.px(28)
        y = h - self.px(13)
        if not ids:
            self._prog_ids = [
                c.create_rectangle(0, 0, 0, 0, fill="#1e2739", outline=""),
                c.create_rectangle(0, 0, 0, 0, fill=ACCENT, outline=""),
            ]
            ids = self._prog_ids
        c.coords(ids[0], pad, y, w - pad, y + self.px(3))
        c.coords(ids[1], pad, y,
                 pad + (w - 2 * pad) * max(0.02, min(1.0, self._prog)),
                 y + self.px(3))

    def _draw_list_card(self):
        c = self.list_card
        if not hasattr(self, "_tree_win"):
            return
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        # 只清自绘内容，保留内嵌的 Treeview / 滚动条窗口（避免每次重建导致闪烁和卡顿）
        c.delete("card")
        c.delete("label")
        draw_card(c, "card", 0, 0, w, h, self.px(16), CARD_SOLID, None, BORDER)

        n = len(self.fonts)
        title = "检测结果" if not n else f"检测结果 · {n} 种字体"
        c.create_text(self.px(20), self.px(25), text=title, anchor="w",
                      fill=TEXT, font=self.fmt(10, "bold"), tags="label")
        if n:
            miss = sum(1 for f in self.fonts
                       if not f["installed"] and f["category"] != "system")
            c.create_text(w - self.px(20), self.px(25),
                          text=(f"{miss} 种缺失" if miss else "全部可用"),
                          anchor="e", fill=(FAIL if miss else OK),
                          font=self.fmt(9, "bold"), tags="label")

        if not self.fonts:
            c.create_text(w // 2, h // 2 - self.px(6), text="尚无检测结果",
                          fill=DIM2, font=self.fmt(11, "bold"), tags="label")
            c.create_text(w // 2, h // 2 + self.px(18),
                          text=DROP_TIP,
                          fill=DIM2, font=self.fmt(8), tags="label")
            for iid in (getattr(self, "_tree_win", None), getattr(self, "_vsb_win", None)):
                if iid:
                    try:
                        c.delete(iid)
                    except Exception:
                        pass
            self._tree_win = None
            self._vsb_win = None
            return

        top = self.px(46)
        tw = max(self.px(120), w - self.px(40))
        th = max(self.px(40), h - top - self.px(12))
        try:
            alive = (self._tree_win is not None and c.type(self._tree_win) != "")
        except Exception:
            alive = False
        if not alive:
            self._tree_win = c.create_window(self.px(14), top, anchor="nw",
                                             window=self.tree, width=tw, height=th)
            self._vsb_win = c.create_window(w - self.px(23), top + self.px(2), anchor="nw",
                                            window=self.vsb, height=th - self.px(4))
        else:
            c.coords(self._tree_win, self.px(14), top)
            c.itemconfig(self._tree_win, width=tw, height=th)
            c.coords(self._vsb_win, w - self.px(23), top + self.px(2))
            c.itemconfig(self._vsb_win, height=max(self.px(30), th - self.px(4)))

    def _draw_sum_card(self):
        c = self.sum_card
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        draw_card(c, "bg", 0, 0, w, h, self.px(14), "#121a27", "#0f1520", BORDER)

        s = self.summary
        if not s:
            c.create_text(w // 2, h // 2 - self.px(6),
                          text="等待检测", fill=DIM2, font=self.fmt(10, "bold"))
            c.create_text(w // 2, h // 2 + self.px(14),
                          text=DROP_TIP2,
                          fill=DIM2, font=self.fmt(8))
            return

        eff_missing = s["missing_count"] - s.get("system_missing_count", 0)
        chips = [
            (f"共 {s['total']} 种", TEXT),
            (f"已装 {s['installed_count']}", OK),
            (f"缺失 {eff_missing}", FAIL if eff_missing else DIM),
            (f"可自动下载 {len(s['opensource_missing'])}", ACCENT),
            (f"需手动获取 {len(s['free_missing']) + len(s['commercial_missing'])}", WARN),
        ]
        pad = self.px(14)
        gap = self.px(10)
        cw = (w - 2 * pad - gap * (len(chips) - 1)) / len(chips)
        cy = self.px(24)
        for i, (txt, color) in enumerate(chips):
            x = pad + i * (cw + gap)
            draw_card(c, "chip", int(x), cy - self.px(13), int(cw), self.px(28),
                      self.px(14), "#1a2332", None, BORDER,
                      bg=_local_bg("#121a27", "#0f1520", cy, h))
            c.create_oval(int(x + self.px(12)), cy - self.px(2),
                          int(x + self.px(18)), cy + self.px(4), fill=color, outline="")
            c.create_text(int(x + self.px(26)), cy, text=txt, anchor="w",
                          fill=color, font=self.fmt(9, "bold"))

        info = (f"本机已安装 {s['installed_total']} 款字体"
                f"　·　文档内嵌 {s['embedded_count']} 款"
                f"　·　耗时 {self._detect_seconds:.2f}s")
        t = s.get("timing") or {}
        if t:
            info += (f"（解析 {t.get('parse', 0):.2f}s · "
                     f"索引 {t.get('index', 0):.2f}s · "
                     f"比对 {t.get('match', 0):.2f}s）")
        sm = s.get("system_missing_count", 0)
        if sm:
            info += f"　·　另有 {sm} 款系统语言字体缺失（不影响中文版式）"
        if s.get("notes"):
            info = s["notes"][0]
        c.create_text(w // 2, self.px(58), text=info, fill=DIM2, font=self.fmt(8))

    # =========================================================
    # 功能 3：关于（开发者信息）
    # =========================================================
    def _build_about_page(self):
        self.about_card = tk.Canvas(self.about_page, bg=BG, highlightthickness=0, bd=0)
        self.about_card.pack(fill="both", expand=True,
                             padx=self.px(34), pady=(0, self.px(18)))
        self.about_card.bind("<Configure>", lambda e: self._later("about", self._draw_about_card))
        self.about_card.bind("<Motion>", self._about_motion)
        self.about_card.bind("<Leave>", self._about_leave)
        self.about_card.bind("<Button-1>", self._about_click)

    def _draw_about_card(self):
        if self.current_mode != "about":
            self._dirty.add("about")
            return
        c = self.about_card
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)

        pad = self.px(30)
        col_w = min(w - 2 * pad, self.px(820))
        x0 = (w - col_w) // 2

        head_h = self.px(92)
        row_h = self.px(32)
        info_h = self.px(22) + row_h * len(app_info.ABOUT_LINES)
        feat_h = self.px(96)
        test_h = self.px(70)
        tips_n = 4
        tips_h = self.px(46) + self.px(19) * tips_n + self.px(8)
        foot_h = self.px(26)
        gap = self.px(18)
        total = (head_h + gap + info_h + gap + feat_h + gap + test_h + gap
                 + tips_h + gap + foot_h + self.px(26))
        y = max(self.px(22), (h - total) // 2)

        # ---------- 头部 ----------
        lsize = self.px(72)
        ph = self.logo_photo(lsize)
        if ph is not None:
            c.create_image(x0, y, image=ph, anchor="nw")
        else:
            draw_card(c, "logo", x0, y, lsize, lsize, self.px(15), ACCENT, ACCENT_D, None,
                      bg=_local_bg(CARD_TOP, CARD_BOT, y + lsize // 2, h))
            c.create_text(x0 + lsize // 2, y + lsize // 2, text="P",
                          fill="#06251b", font=self.fmt(23, "bold"))
        tx = x0 + lsize + self.px(16)
        c.create_text(tx, y + self.px(20), text=app_info.APP_NAME, anchor="w",
                      fill=TEXT, font=self.fmt(19, "bold"))
        c.create_text(tx, y + self.px(48), text=app_info.SUBTITLE, anchor="w",
                      fill=DIM, font=self.fmt(9))

        bw, bh = self.px(78), self.px(26)
        bx = x0 + col_w - bw
        ty = y + self.px(6)
        draw_card(c, "tag2", bx, ty, bw, bh, bh // 2, "#1d5c47", "#164a3a", ACCENT_D,
                  bg=_local_bg(CARD_TOP, CARD_BOT, ty + bh // 2, h))
        c.create_text(bx + bw // 2, ty + bh // 2, text=app_info.TAGLINE,
                      fill=ACCENT, font=self.fmt(9, "bold"))
        bw2 = self.px(70)
        bx2 = bx - self.px(8) - bw2
        draw_card(c, "tag1", bx2, ty, bw2, bh, bh // 2, "#1a2231", None, BORDER,
                  bg=_local_bg(CARD_TOP, CARD_BOT, ty + bh // 2, h))
        c.create_text(bx2 + bw2 // 2, ty + bh // 2, text="v" + app_info.VERSION,
                      fill=DIM, font=self.fmt(9, "bold"))

        # ---------- 开发者信息 ----------
        y += head_h + gap
        draw_card(c, "info", x0, y, col_w, info_h, self.px(14), "#121a27", None, BORDER,
                  bg=_local_bg(CARD_TOP, CARD_BOT, y + info_h // 2, h))
        ly = y + self.px(22) + row_h // 2
        for label, value in app_info.ABOUT_LINES:
            c.create_text(x0 + self.px(22), ly, text=label, anchor="w",
                          fill=DIM2, font=self.fmt(9))
            color = ACCENT if label == "开发者" else TEXT
            size = 12 if label == "开发者" else 10
            c.create_text(x0 + self.px(112), ly, text=value, anchor="w",
                          fill=color, font=self.fmt(size, "bold" if label == "开发者" else "normal"))
            ly += row_h

        # ---------- 功能说明 ----------
        y += info_h + gap
        fgap = self.px(12)
        fw = (col_w - 2 * fgap) // 3
        feats = [
            ("①", "字体检测", "解析版式 / 母版 / 主题 / 图表 / 内嵌字体，缺失的开源字体可一键下载安装"),
            ("②", "页面图片化", "每页导出高清图再插回新 PPT，版式彻底固定，源文件只读不改"),
            ("③", "自动保存", "按间隔自动保存并另存副本，%s 均支持" % APP_FAMILY),
        ]
        for i, (num, name, desc) in enumerate(feats):
            fx = x0 + i * (fw + fgap)
            draw_card(c, f"feat{i}", fx, y, fw, feat_h, self.px(14), "#141c2a", None, BORDER,
                      bg=_local_bg(CARD_TOP, CARD_BOT, y + feat_h // 2, h))
            bsize = self.px(26)
            draw_card(c, f"fnum{i}", fx + self.px(16), y + self.px(16), bsize, bsize,
                      self.px(9), "#1d5c47", "#164a3a", ACCENT_D, bg="#141c2a")
            c.create_text(fx + self.px(16) + bsize // 2, y + self.px(16) + bsize // 2,
                          text=num, fill=ACCENT, font=self.fmt(10, "bold"))
            c.create_text(fx + self.px(50), y + self.px(29), text=name, anchor="w",
                          fill=TEXT, font=self.fmt(11, "bold"))
            c.create_text(fx + self.px(16), y + self.px(52), text=desc, anchor="nw",
                          fill=DIM, font=self.fmt(8),
                          width=fw - self.px(30))

        # ---------- 功能自检（最显眼的入口，报错时先点这里） ----------
        y += feat_h + gap
        self._draw_selftest_entry(c, x0, y, col_w, test_h)

        # ---------- 使用提示 ----------
        y += test_h + gap
        draw_card(c, "tips", x0, y, col_w, tips_h, self.px(14), "#121a27", None, BORDER,
                  bg=_local_bg(CARD_TOP, CARD_BOT, y + tips_h // 2, h))
        c.create_text(x0 + self.px(22), y + self.px(22), text="使用提示", anchor="w",
                      fill=DIM, font=self.fmt(9, "bold"))
        tips = [
            "检测过程只读，不会改动你的 PPT；本机字体索引会缓存到本地，第二次检测几乎瞬间完成",
            "页面图片化会另存为新文件，源文件保持原样；图片化后的页面不能再编辑",
            "支持 .pptx / .pptm / .ppt / .pps / .ppsx / .potx，无需安装 Office 即可检测字体",
            "自动保存只处理当前打开的文档，「定时备份」默认存到「文档%sPPT自动备份」，可在设置里改"
            % os.sep,
        ]
        for i, tip in enumerate(tips):
            yy = y + self.px(46) + i * self.px(19)
            c.create_text(x0 + self.px(22), yy, text="·", anchor="w",
                          fill=ACCENT_D, font=self.fmt(9, "bold"))
            c.create_text(x0 + self.px(34), yy, text=tip, anchor="w",
                          fill=DIM2, font=self.fmt(8),
                          width=col_w - self.px(56))

        # ---------- 底部 ----------
        c.create_text(w // 2, y + tips_h + gap + foot_h // 2,
                      text="公益免费软件 · 禁止倒卖与二次收费 · 转载或分发请保留作者与版本信息",
                      fill=DIM2, font=self.fmt(8))

    def _draw_selftest_entry(self, c, x0, y, w, h):
        """关于页上的「一键功能自检」：报错时最该被点到的那个按钮。"""
        hot = bool(getattr(self, "_st_hot", False))
        top = "#1d4438" if hot else "#173a2f"
        bot = "#132e26" if hot else "#102720"
        border = ACCENT if hot else ACCENT_D
        card_bg = _local_bg(CARD_TOP, CARD_BOT, y + h // 2, c.winfo_height())
        draw_card(c, "stbtn", x0, y, w, h, self.px(14), top, bot, border, bg=card_bg)

        d = self.px(40)
        iy = y + (h - d) // 2
        draw_card(c, "stbtn", x0 + self.px(18), iy, d, d, d // 2,
                  ACCENT, ACCENT_D, None, bg=top)
        c.create_text(x0 + self.px(18) + d // 2, iy + d // 2, text="◎",
                      fill="#06251b", font=self.fmt(15, "bold"), tags=("stbtn",))

        tx = x0 + self.px(18) + d + self.px(14)
        c.create_text(tx, y + self.px(20),
                      text="一键功能自检 · 内置 3 份测试 PPT",
                      anchor="w", fill=TEXT, font=self.fmt(12, "bold"),
                      tags=("stbtn",))
        c.create_text(tx, y + self.px(45),
                      text="报错时先点这里：资源释放 → 路径 → 字体检测 → 自动保存 → 图片化，逐项给结论",
                      anchor="w", fill="#9fe3cd" if hot else DIM,
                      font=self.fmt(8), tags=("stbtn",))

        bw, bh = self.px(112), self.px(34)
        bx = x0 + w - bw - self.px(18)
        by = y + (h - bh) // 2
        draw_card(c, "stbtn", bx, by, bw, bh, bh // 2,
                  "#3ce3aa" if hot else ACCENT, ACCENT_D, None, bg=top)
        c.create_text(bx + bw // 2, by + bh // 2, text="开始检测  →",
                      fill="#06251b", font=self.fmt(9, "bold"), tags=("stbtn",))

        self._st_rect = (x0, y, w, h)

    def _st_hit(self, x, y):
        r = getattr(self, "_st_rect", None)
        if not r:
            return False
        return r[0] <= x <= r[0] + r[2] and r[1] <= y <= r[1] + r[3]

    def _about_motion(self, event):
        if self.current_mode != "about":
            return
        hot = self._st_hit(event.x, event.y)
        if hot != bool(getattr(self, "_st_hot", False)):
            self._st_hot = hot
            try:
                self.about_card.configure(cursor="hand2" if hot else "")
            except Exception:
                pass
            self._draw_about_card()

    def _about_leave(self, event):
        if getattr(self, "_st_hot", False):
            self._st_hot = False
            try:
                self.about_card.configure(cursor="")
            except Exception:
                pass
            self._draw_about_card()

    def _about_click(self, event):
        if self.current_mode == "about" and self._st_hit(event.x, event.y):
            self.start_selftest()

    def start_selftest(self):
        """打开功能自检窗口（单实例，重复点只前置）。"""
        dlg = getattr(self, "_st_dlg", None)
        if dlg is None:
            dlg = SelfTestDialog(self)
            self._st_dlg = dlg
        dlg.open()

    # =========================================================
    # 功能 3：自动保存 / 定时备份
    # =========================================================
    def _build_autosave_page(self):
        page = self.autosave_page

        self.as_status_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                        height=self.px(128))
        self.as_status_card.pack(fill="x", padx=self.px(34), pady=(0, self.px(12)))
        self.as_status_card.bind("<Configure>",
                                 lambda e: self._later("asst", self._draw_as_status))

        self.as_set_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                     height=self.px(272))
        self.as_set_card.pack(fill="x", padx=self.px(34), pady=(0, self.px(12)))
        self.as_set_card.bind("<Configure>",
                              lambda e: self._later("asset", self._draw_as_settings))
        self.as_set_card.bind("<Button-1>", self._as_click)
        self.as_set_card.bind("<Motion>", self._as_motion)
        self.as_set_card.bind("<Leave>",
                              lambda e: self.as_set_card.configure(cursor=""))

        # 按钮先按 bottom 占位，日志再吃掉剩余空间：窗口再小也不会把按钮挤没
        btns = tk.Frame(page, bg=BG)
        btns.pack(side="bottom", pady=(self.px(14), self.px(16)))
        self.btn_as_start = self._pill(btns, "开始监听", self.as_start, "primary", 128, 42)
        self.btn_as_start.pack(side="left", padx=self.px(6))
        self.btn_as_stop = self._pill(btns, "停止", self.as_stop, "ghost", 88, 42)
        self.btn_as_stop.pack(side="left", padx=self.px(6))
        self.btn_as_now = self._pill(btns, "立即备份", self.as_backup_now, "ghost", 114, 42)
        self.btn_as_now.pack(side="left", padx=self.px(6))
        self.btn_as_open = self._pill(btns, "打开备份目录", self.as_open_dir, "ghost", 148, 42)
        self.btn_as_open.pack(side="left", padx=self.px(6))

        self.as_log_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                     height=self.px(140))
        self.as_log_card.pack(fill="both", expand=True, padx=self.px(34))
        self.as_log_card.bind("<Configure>",
                              lambda e: self._later("aslog", self._draw_as_log))

    # ---------------- 绘制 ----------------
    def _draw_as_all(self):
        self._draw_as_status()
        self._draw_as_settings()
        self._draw_as_log()

    def _draw_as_status(self):
        if self.current_mode != "autosave":
            self._dirty.add("autosave")
            return
        c = self.as_status_card
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)
        st, _ = self.as_engine.snapshot()

        running = bool(st.get("running"))
        attached = bool(st.get("attached"))
        docs = st.get("doc_names") or []

        c.create_text(self.px(22), self.px(20), text="监听状态", anchor="w",
                      fill=DIM, font=self.fmt(9, "bold"))

        if running and attached:
            label, color = "监听中", ACCENT
        elif running:
            label, color = "监听中 · 等待文档", "#f0b350"
        else:
            label, color = "未开始", DIM

        cy = self.px(52)
        c.create_oval(self.px(22), cy - self.px(5), self.px(32), cy + self.px(5),
                      fill=color, outline="")
        c.create_text(self.px(42), cy, text=label, anchor="w", fill=TEXT,
                      font=self.fmt(15, "bold"))

        # 右上：连接的是哪个程序、几个文档
        if attached:
            side = "%s · %d 个文档" % (st.get("app") or "演示程序", len(docs))
        elif running:
            side = "未检测到正在运行的 " + APP_FAMILY
        else:
            side = "点「开始监听」后在后台自动保存"
        c.create_text(w - self.px(22), cy, text=self._ellipsis(side, 52), anchor="e",
                      fill=(DIM if attached else DIM2), font=self.fmt(9))

        # 当前文档
        if docs:
            names = "、".join(docs[:4]) + ("等" if len(docs) > 4 else "")
            doc_txt = "当前文档：" + names
            doc_color = TEXT
        else:
            doc_txt = ("提示：先打开要编辑的 PPT，再点开始监听" if not running
                       else "等待打开演示文稿…")
            doc_color = DIM2
        c.create_text(self.px(22), self.px(82), text=self._ellipsis(doc_txt, 78),
                      anchor="w", fill=doc_color, font=self.fmt(9))

        # 上次动作 + 下次时间
        def hhmmss(ts):
            return time.strftime("%H:%M:%S", time.localtime(ts)) if ts else "—"

        cfg = self.as_cfg
        next_bits = []
        if running and cfg.get("save_enabled"):
            nxt = self.as_engine.next_save_at()
            if nxt:
                next_bits.append("下次保存 " + time.strftime("%H:%M:%S", time.localtime(nxt)))
        if running and cfg.get("backup_enabled"):
            nxt = self.as_engine.next_backup_at()
            if nxt:
                next_bits.append("下次备份 " + time.strftime("%H:%M:%S", time.localtime(nxt)))
        tail = ("　·　" + "　".join(next_bits)) if next_bits else ""
        c.create_text(self.px(22), self.px(106),
                      text="上次保存 %s　上次备份 %s%s" % (
                          hhmmss(st.get("last_save")), hhmmss(st.get("last_backup")), tail),
                      anchor="w", fill=DIM2, font=self.fmt(8))

    def _draw_as_settings(self):
        c = self.as_set_card
        c.delete("all")
        self._as_hits = []
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)
        cfg = self.as_cfg

        c.create_text(self.px(22), self.px(24), text="保存与备份设置", anchor="w",
                      fill=TEXT, font=self.fmt(10, "bold"))
        c.create_text(w - self.px(22), self.px(24),
                      text="间隔可自由设定 · %s 通用" % APP_FAMILY,
                      anchor="e", fill=DIM2, font=self.fmt(8))

        # ---- 第一行：编辑时自动保存 ----
        y1 = self.px(64)
        self._as_row_label(c, "编辑时自动保存", y1)
        self._as_switch(c, self.px(150), y1, bool(cfg.get("save_enabled")), "toggle_save", h)
        c.create_text(self.px(208), y1, text="每隔", anchor="w", fill=DIM2, font=self.fmt(9))
        self._as_interval_chips(c, w, h, self.px(246), y1, "save_interval_min",
                                SAVE_PRESETS, bool(cfg.get("save_enabled")))
        c.create_text(w - self.px(22), y1,
                      text="当前 " + autosave.interval_text(cfg.get("save_interval_min")),
                      anchor="e", fill=DIM, font=self.fmt(9))

        # ---- 第二行：定时备份 ----
        y2 = self.px(120)
        self._as_row_label(c, "定时备份", y2)
        self._as_switch(c, self.px(150), y2, bool(cfg.get("backup_enabled")), "toggle_backup", h)
        c.create_text(self.px(208), y2, text="每隔", anchor="w", fill=DIM2, font=self.fmt(9))
        x = self._as_interval_chips(c, w, h, self.px(246), y2, "backup_interval_min",
                                    BACKUP_PRESETS, bool(cfg.get("backup_enabled")))
        c.create_text(x + self.px(14), y2, text="备份到", anchor="w", fill=DIM2, font=self.fmt(9))
        bw, bh = self.px(96), self.px(28)
        bx, by2 = x + self.px(60), y2 - bh // 2
        draw_card(c, "asdir", bx, by2, bw, bh, bh // 2, BLUE, BLUE_D, None,
                  bg=_local_bg(CARD_TOP, CARD_BOT, y2, h))
        c.create_text(bx + bw // 2, y2, text="选择目录", fill="white", font=self.fmt(9, "bold"))
        self._as_hits.append((bx, by2, bw, bh, ("pick_dir", None)))

        # ---- 第三行：目录与保留份数 ----
        y3 = self.px(156)
        folder = cfg.get("backup_dir") or ""
        if folder:
            dirc = self._ellipsis(folder, 62)
            color = TEXT
        else:
            dirc = "未指定（默认存到「文档\\PPT自动备份」）"
            color = DIM2
        c.create_text(self.px(22), y3, text="备份位置：", anchor="w", fill=DIM2, font=self.fmt(9))
        c.create_text(self.px(84), y3, text=dirc, anchor="w", fill=color, font=self.fmt(9))
        c.create_text(w - self.px(22), y3,
                      text="每个文档保留最近 %d 份" % int(cfg.get("keep_copies") or 60),
                      anchor="e", fill=DIM2, font=self.fmt(8))

        # ---- 第四行：开机自启动 ----
        y4 = self.px(198)
        st = self._as_auto_state or autostart.state()
        self._as_row_label(c, "开机自启动", y4)
        self._as_switch(c, self.px(150), y4, bool(st.get("enabled")), "toggle_auto", h)
        hint = ("已加入启动项：登录 %s 后自动在后台开始监听" % LOGIN_WORD
                if st.get("enabled") else
                "开启后，开机自动在后台开始监听（%s）" % AUTOSTART_HINT)
        c.create_text(self.px(208), y4, text=self._ellipsis(hint, 62), anchor="w",
                      fill=(ACCENT_D if st.get("enabled") else DIM2), font=self.fmt(8))

        # ---- 说明 ----
        c.create_text(self.px(22), self.px(234),
                      text="说明：「自动保存」等同你按 %s，只写回原文件；「定时备份」另存独立副本，不改动原文件。" % SAVE_KEY,
                      anchor="w", fill=DIM2, font=self.fmt(8))
        c.create_text(self.px(22), self.px(252),
                      text="只处理当前打开的文档；只读文档与从未保存过的新文档会自动跳过，并在下方记录原因。",
                      anchor="w", fill=DIM2, font=self.fmt(8))

    def _as_row_label(self, c, text, cy):
        c.create_text(self.px(22), cy, text=text, anchor="w", fill=DIM, font=self.fmt(9))

    def _as_switch(self, c, x, cy, on, action, card_h):
        w, hh = self.px(44), self.px(24)
        y = cy - hh // 2
        draw_card(c, "sw_" + action, x, y, w, hh, hh // 2,
                  "#1d5c47" if on else "#1a2231",
                  "#164a3a" if on else "#151c29",
                  ACCENT_D if on else BORDER,
                  bg=_local_bg(CARD_TOP, CARD_BOT, cy, card_h))
        d = hh - self.px(6)
        kx = x + (w - d - self.px(3)) if on else x + self.px(3)
        c.create_oval(kx, y + self.px(3), kx + d, y + self.px(3) + d,
                      fill=ACCENT if on else "#5b6579", outline="")
        self._as_hits.append((x, y, w, hh, (action, None)))

    def _as_interval_chips(self, c, w, h, x0, cy, key, presets, enabled):
        """画一排间隔选项 + 自定义；返回结束时的 x 坐标"""
        cfg = self.as_cfg
        cur = autosave.clamp_interval(cfg.get(key))
        cw, ch = self.px(46), self.px(28)
        x = x0
        y = cy - ch // 2
        for val in presets:
            on = enabled and abs(cur - val) < 0.001
            draw_card(c, "chip_%s_%s" % (key, val), x, y, cw, ch, ch // 2,
                      "#1d5c47" if on else "#1a2231",
                      "#164a3a" if on else "#151c29",
                      ACCENT_D if on else BORDER,
                      bg=_local_bg(CARD_TOP, CARD_BOT, cy, h))
            c.create_text(x + cw // 2, cy, text=str(val),
                          fill=ACCENT if on else DIM,
                          font=self.fmt(9, "bold" if on else "normal"))
            self._as_hits.append((x, y, cw, ch, (key, val)))
            x += cw + self.px(6)

        custom_on = enabled and all(abs(cur - v) > 0.001 for v in presets)
        cw2 = self.px(78)
        draw_card(c, "chip_%s_custom" % key, x, y, cw2, ch, ch // 2,
                  "#1d5c47" if custom_on else "#1a2231",
                  "#164a3a" if custom_on else "#151c29",
                  ACCENT_D if custom_on else BORDER,
                  bg=_local_bg(CARD_TOP, CARD_BOT, cy, h))
        c.create_text(x + cw2 // 2, cy,
                      text=("每 %s" % autosave.interval_text(cur)) if custom_on else "自定义",
                      fill=ACCENT if custom_on else DIM,
                      font=self.fmt(8, "bold" if custom_on else "normal"))
        self._as_hits.append((x, y, cw2, ch, (key, "custom")))
        return x + cw2

    def _draw_as_log(self):
        c = self.as_log_card
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 10 or h < 10:
            return
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)
        c.create_text(self.px(22), self.px(22), text="最近记录", anchor="w",
                      fill=TEXT, font=self.fmt(10, "bold"))
        _, logs = self.as_engine.snapshot()
        c.create_text(w - self.px(22), self.px(22), text="共 %d 条" % len(logs),
                      anchor="e", fill=DIM2, font=self.fmt(8))

        line_h = self.px(19)
        top = self.px(46)
        max_n = max(1, (h - top - self.px(12)) // line_h)
        rows = logs[-max_n:]
        if not rows:
            c.create_text(self.px(22), top + line_h // 2,
                          text="暂无记录。点「开始监听」后，每次保存与备份都会记在这里。",
                          anchor="w", fill=DIM2, font=self.fmt(9))
            return
        colors = {"ok": ACCENT, "warn": "#f0b350", "error": "#ff7b72",
                  "info": DIM, "note": DIM2}
        for i, (ts, kind, msg) in enumerate(rows):
            y = top + i * line_h + line_h // 2
            c.create_text(self.px(22), y,
                          text=time.strftime("%H:%M:%S", time.localtime(ts)),
                          anchor="w", fill=DIM2, font=self.fmt(8))
            c.create_text(self.px(90), y, text=self._ellipsis(msg, 100), anchor="w",
                          fill=colors.get(kind, DIM), font=self.fmt(9))

    @staticmethod
    def _ellipsis(text, limit):
        text = text or ""
        return text if len(text) <= limit else text[:limit - 1] + "…"

    # ---------------- 交互 ----------------
    def _as_hit_at(self, x, y):
        for hx, hy, hw, hh, action in self._as_hits:
            if hx <= x <= hx + hw and hy <= y <= hy + hh:
                return action
        return None

    def _as_motion(self, event):
        self.as_set_card.configure(
            cursor="hand2" if self._as_hit_at(event.x, event.y) else "")

    def _as_click(self, event):
        action = self._as_hit_at(event.x, event.y)
        if not action:
            return
        kind, val = action
        cfg = self.as_cfg

        if kind in ("toggle_save", "toggle_backup"):
            key = "save_enabled" if kind == "toggle_save" else "backup_enabled"
            cfg[key] = not bool(cfg.get(key))
            autosave.save_config(cfg)
            if cfg[key] and not self.as_engine.running:
                self.as_start()          # 打开开关即开始监听
                return
            self.as_engine.reset_timers()
            self._draw_as_settings()
            self._sync_as_buttons()
            return

        if kind in ("save_interval_min", "backup_interval_min"):
            if val == "custom":
                self._as_custom_interval(kind)
            else:
                cfg[kind] = float(val)
                autosave.save_config(cfg)
                self.as_engine.reset_timers()
                self._draw_as_settings()
            return

        if kind == "pick_dir":
            self._as_pick_dir()
            return

        if kind == "toggle_auto":
            self._as_toggle_autostart()

    def _as_custom_interval(self, key):
        cfg = self.as_cfg
        cur = autosave.clamp_interval(cfg.get(key))
        title = "自动保存间隔" if key == "save_interval_min" else "定时备份间隔"
        try:
            v = simpledialog.askinteger(
                title, "每隔多少分钟执行一次？（1–720，输入 1 表示每分钟）",
                parent=self.root, initialvalue=int(round(cur)),
                minvalue=1, maxvalue=720)
        except Exception:
            v = None
        if not v:
            return
        cfg[key] = float(v)
        autosave.save_config(cfg)
        self.as_engine.reset_timers()
        self._draw_as_settings()

    def _as_pick_dir(self):
        cfg = self.as_cfg
        cur = cfg.get("backup_dir") or autosave.default_backup_dir()
        init = paths.safe_initialdir(cur)
        d = filedialog.askdirectory(title="选择备份存放目录", initialdir=init)
        if not d:
            return
        d = paths.abspath(d)
        if not paths.is_dir(d):
            if not paths.ensure_dir(d):
                messagebox.showerror("目录不可用",
                                     "无法创建或访问这个目录：\n%s" % d)
                return
        if not paths.is_writable(d):
            messagebox.showerror("目录不可写", "这个目录不能写入：\n%s\n请换一个位置。" % d)
            return
        cfg["backup_dir"] = d
        autosave.save_config(cfg)
        self._draw_as_settings()

    def _as_toggle_autostart(self):
        st = autostart.state()
        enabling = not st.get("enabled")
        if enabling:
            ok, err = autostart.enable()
            what = "已加入开机启动项"
        else:
            ok, err = autostart.disable()
            what = "已关闭开机自启动"
        self._as_auto_state = autostart.state()
        if not ok:
            self._draw_as_settings()
            messagebox.showerror("开机自启动", "设置失败：\n%s" % err)
            return
        # 打开自启动就顺手开始监听：否则注册表加了、程序却在后台什么都不做
        if enabling and not self.as_engine.running:
            self.as_start()
        self._draw_as_settings()
        self.as_engine.log("info", "%s（%s）" % (what, self._as_auto_state.get("command", "")))
        self._draw_as_log()

    def _sync_as_buttons(self):
        running = self.as_engine.running
        for btn, en in ((self.btn_as_start, not running),
                        (self.btn_as_stop, running),
                        (self.btn_as_now, True),
                        (self.btn_as_open, True)):
            try:
                btn.set_enabled(en)
            except Exception:
                pass

    # ---------------- 动作 ----------------
    def as_start(self):
        cfg = self.as_cfg
        if not cfg.get("save_enabled") and not cfg.get("backup_enabled"):
            cfg["save_enabled"] = True      # 至少让一件事能跑起来
        if cfg.get("backup_enabled") and not cfg.get("backup_dir"):
            cfg["backup_dir"] = autosave.default_backup_dir()
        cfg["enabled"] = True
        autosave.save_config(cfg)
        self.as_engine.start()
        self._as_sig = None
        self._draw_as_all()
        self._sync_as_buttons()

    def as_stop(self):
        self.as_cfg["enabled"] = False
        autosave.save_config(self.as_cfg)
        self.as_engine.stop()
        self._as_sig = None
        self._draw_as_all()
        self._sync_as_buttons()

    def as_backup_now(self):
        if not self.as_engine.running:
            self.as_start()
        self.as_engine.force_backup()
        self._draw_as_log()

    def as_open_dir(self):
        folder = self.as_cfg.get("backup_dir") or autosave.default_backup_dir()
        if not paths.is_dir(folder):
            if not paths.ensure_dir(folder):
                messagebox.showerror("目录不可用", "无法创建或访问：\n%s" % folder)
                return
        try:
            paths.open_path(folder)
        except Exception as exc:
            messagebox.showerror("打开失败", "%s\n%s" % (folder, exc))

    # ---------------- 定时刷新 ----------------
    def _as_refresh(self):
        try:
            if self.current_mode == "autosave":
                st, logs = self.as_engine.snapshot()
                sig = (st.get("running"), st.get("attached"), st.get("app"),
                       st.get("docs"), tuple(st.get("doc_names") or ()),
                       st.get("last_save"), st.get("last_backup"),
                       len(logs), st.get("error"))
                if sig != self._as_sig:
                    self._as_sig = sig
                    self._draw_as_status()
                    self._draw_as_log()
                    self._sync_as_buttons()
        except Exception:
            pass
        self.root.after(800, self._as_refresh)

    def enter_background_mode(self):
        """--autostart 用：静默开始监听并最小化到任务栏。

        开机启动项在 = 用户希望开机就工作，所以即使上次是手动停止的，
        这里也按启动项的状态把监听起来（「停止」只影响当次运行）。
        """
        cfg = self.as_cfg
        wants = cfg.get("enabled") or autostart.is_enabled()
        if wants and (cfg.get("save_enabled") or cfg.get("backup_enabled")):
            self.as_start()
        self.root.iconify()

    def _on_close(self):
        """关窗口前把监听线程收干净；正在监听则保留状态，下次开机继续。"""
        try:
            if self.as_engine.running:
                self.as_cfg["enabled"] = True
                autosave.save_config(self.as_cfg)
                self.as_engine.stop()
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    # =========================================================
    # 功能 2：页面图片化页
    # =========================================================
    def _build_flatten_page(self):
        page = self.flatten_page

        self.fdrop_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                    height=self.px(108))
        self.fdrop_card.pack(fill="x", padx=self.px(34), pady=(0, self.px(14)))
        self.fdrop_card.bind("<Configure>", lambda e: self._later("fdrop", self._draw_fdrop_card))

        self.fset_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                   height=self.px(168))
        self.fset_card.pack(fill="x", padx=self.px(34), pady=(0, self.px(14)))
        self.fset_card.bind("<Configure>", lambda e: self._later("fset", self._draw_fset_card))
        # 命中区域在画布层统一判定：画布上的文字/贴片会盖住下层 item，
        # 用 tag_bind 只在点中卡片本体时有效，点文字就没反应
        self._fset_hits = []
        self.fset_card.bind("<Button-1>", self._fset_click)
        self.fset_card.bind("<Motion>", self._fset_motion)
        self.fset_card.bind("<Leave>", lambda e: self.fset_card.configure(cursor=""))

        self.fprog_card = tk.Canvas(page, bg=BG, highlightthickness=0, bd=0,
                                    height=self.px(150))
        self.fprog_card.pack(fill="both", expand=True, padx=self.px(34))
        self.fprog_card.bind("<Configure>", lambda e: self._later("fprog", self._draw_fprog_card))

        btns = tk.Frame(page, bg=BG)
        btns.pack(pady=(self.px(16), self.px(18)))
        self.btn_fchoose = self._pill(btns, "选择文件", self.fchoose_file, "ghost", 118, 42)
        self.btn_fchoose.pack(side="left", padx=self.px(6))
        self.btn_frun = self._pill(btns, "开始图片化", self.frun, "primary", 152, 42)
        self.btn_frun.pack(side="left", padx=self.px(6))
        self.btn_fclear = self._pill(btns, "清空", self.fclear, "ghost", 96, 42)
        self.btn_fclear.pack(side="left", padx=self.px(6))

    def _draw_fdrop_card(self):
        if self.current_mode != "flatten":
            self._dirty.add("flatten")
            return
        c = self.fdrop_card
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)
        cx = w // 2
        name = getattr(self, "_floaded_name", None)
        c.create_text(cx, h // 2 - self.px(16), text=(name or DROP_TITLE),
                      fill=TEXT, font=self.fmt(14, "bold"))
        c.create_text(cx, h // 2 + self.px(8),
                      text="每页导出为高清图片，再插回新 PPT，彻底固定版式",
                      fill=DIM, font=self.fmt(9))
        c.create_text(cx, h // 2 + self.px(28),
                      text="源文件只读不改 · 输出为新文件", fill=DIM2, font=self.fmt(8))

    def _draw_fset_card(self):
        c = self.fset_card
        c.delete("all")
        self._fset_hits = []
        w, h = c.winfo_width(), c.winfo_height()
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_TOP, CARD_BOT, BORDER)

        c.create_text(self.px(22), self.px(24), text="导出设置", anchor="w",
                      fill=TEXT, font=self.fmt(10, "bold"))

        # 分辨率
        c.create_text(self.px(22), self.px(58), text="图片分辨率", anchor="w",
                      fill=DIM, font=self.fmt(9))
        x = self.px(120)
        for val, label in DPI_OPTIONS:
            on = (self.dpi_var.get() == val)
            bw, bh = self.px(92), self.px(30)
            px_, py_ = x, self.px(43)
            draw_card(c, f"dpi{val}", px_, py_, bw, bh, bh // 2,
                      ("#1d5c47" if on else "#1a2231"),
                      ("#164a3a" if on else "#151c29"),
                      (ACCENT_D if on else BORDER),
                      bg=_local_bg(CARD_TOP, CARD_BOT, self.px(58), h))
            c.create_text(px_ + bw // 2, self.px(58), text=label,
                          fill=(ACCENT if on else DIM),
                          font=self.fmt(9, "bold" if on else "normal"))
            self._fset_hits.append((px_, py_, bw, bh, ("dpi", val)))
            x += bw + self.px(8)

        # 输出位置
        c.create_text(self.px(22), self.px(100), text="输出位置", anchor="w",
                      fill=DIM, font=self.fmt(9))
        out = self._fout_dir or "源文件同目录"
        if len(out) > 46:
            out = "…" + out[-44:]
        c.create_text(self.px(120), self.px(100), text=out, anchor="w",
                      fill=TEXT, font=self.fmt(9))

        bw, bh = self.px(104), self.px(32)
        bx = w - self.px(22) - bw
        by = self.px(84)
        draw_card(c, "outbtn", bx, by, bw, bh, bh // 2, BLUE, BLUE_D, None,
                  bg=_local_bg(CARD_TOP, CARD_BOT, by + bh // 2, h))
        c.create_text(bx + bw // 2, by + bh // 2, text="选择目录", fill="white",
                      font=self.fmt(9, "bold"))
        self._fset_hits.append((bx, by, bw, bh, ("outdir", None)))

        c.create_text(self.px(22), self.px(136),
                      text="说明：输出为新的 .pptx，源文件不会被修改；图片化后内容不可编辑（由本机 PowerPoint 渲染）。",
                      anchor="w", fill=DIM2, font=self.fmt(8))

    def _fset_hit_at(self, x, y):
        for hx, hy, hw, hh, action in self._fset_hits:
            if hx <= x <= hx + hw and hy <= y <= hy + hh:
                return action
        return None

    def _fset_motion(self, event):
        self.fset_card.configure(
            cursor="hand2" if self._fset_hit_at(event.x, event.y) else "")

    def _fset_click(self, event):
        if self._flattening:
            return
        action = self._fset_hit_at(event.x, event.y)
        if not action:
            return
        kind, val = action
        if kind == "dpi":
            self._set_dpi(val)
        elif kind == "outdir":
            self.choose_out_dir()

    def _set_dpi(self, val):
        if self._flattening:
            return
        if self.dpi_var.get() == val:
            return
        self.dpi_var.set(val)
        self._draw_fset_card()

    def _draw_fprog_card(self):
        c = self.fprog_card
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        draw_card(c, "bg", 0, 0, w, h, self.px(16), CARD_SOLID, None, BORDER)

        c.create_text(self.px(22), self.px(24), text="转换进度", anchor="w",
                      fill=TEXT, font=self.fmt(10, "bold"))
        c.create_text(w // 2, h // 2 - self.px(6), text=self._fstatus,
                      fill=TEXT, font=self.fmt(12, "bold"))
        if self._fhint:
            c.create_text(w // 2, h // 2 + self.px(18), text=self._fhint,
                          fill=DIM2, font=self.fmt(8))

        pad = self.px(22)
        y = h - self.px(20)
        c.create_rectangle(pad, y, w - pad, y + self.px(4),
                           fill="#1e2739", outline="")
        if self._fprog:
            c.create_rectangle(pad, y, pad + (w - 2 * pad) * max(0.02, self._fprog),
                               y + self.px(4), fill=ACCENT, outline="")

    def _set_fprog(self, status=None, hint=None, ratio=None):
        if status is not None:
            self._fstatus = status
        if hint is not None:
            self._fhint = hint
        if ratio is not None:
            self._fprog = ratio
        self._draw_fprog_card()

    def choose_out_dir(self):
        if self._flattening:
            return
        init = paths.safe_initialdir(
            self._fout_dir or (os.path.dirname(self.files[0]) if self.files else ""))
        kw = {"title": "选择输出目录"}
        if init:
            kw["initialdir"] = init
        d = filedialog.askdirectory(**kw)
        if d:
            self._fout_dir = paths.normalize(d)
            self._draw_fset_card()

    # ---------------- 拖拽 ----------------
    def _setup_drag_drop(self):
        """用原生 DragQueryFileW 接管拖放，路径来自 UTF-16，与代码页无关

        macOS 的 Tk 没有 WM_DROPFILES，原生拖放不可用：直接跳过，
        界面上会提示改用「选择文件」/「载入文件」。
        """
        if MACOS:
            return
        # 允许重复调用：dropfiles 内部按句柄去重，窗口映射后再补一次，
        # 否则只会勾到内层 TkChild，真实拖拽投递到外层窗口就收不到
        if self._hook_tries >= 8:
            return
        self._hook_tries += 1
        try:
            import dropfiles
            n = dropfiles.hook(self.root, self._on_drop_files)
            self._drop_hooked = self._drop_hooked or bool(n)
            dropfiles.log("hooked %d windows" % n)
        except Exception as e:
            self._drop_hooked = False
            try:
                import dropfiles
                dropfiles.log("hook failed: %r" % (e,))
            except Exception:
                pass

    @staticmethod
    def _repair_path(path):
        """把拖进来 / 粘贴进来的路径尽量还原成真实存在的路径。

        覆盖：带引号、file:// 形式、%VAR%、超长路径、8.3 短名、
        被错误解码的中文（UTF-8 当 GBK 解等）、段尾多空格/点、大小写不一致。
        修不好时原样返回，由上层给出可操作的提示，不静默失败。
        """
        try:
            return paths.resolve_existing(path)
        except Exception:
            return path

    def _on_drop_files(self, files):
        # 回调在主线程的消息泵里执行，但仍只入队，保持单一处理路径
        try:
            import dropfiles
            dropfiles.log("drop cb: %r" % (files,))
        except Exception:
            pass
        if self.detecting or self.installing or self._flattening:
            return
        for f in files:
            if isinstance(f, bytes):
                f = self._repair_path(f.decode("utf-8", errors="surrogateescape"))
            else:
                f = self._repair_path(f)
            if f.lower().endswith((".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx")):
                self._event_queue.put(("drop", f))
                break

    def _handle_dropped(self, path):
        if self.detecting or self.installing or self._flattening:
            return
        if self.current_mode == "autosave":
            # 自动保存页不接收文件拖放，避免莫名其妙开始字体检测
            return
        path = self._repair_path(path)
        if not paths.is_file(path):
            reasons = "\n".join("· " + r for r in paths.describe_missing(path))
            messagebox.showerror(
                "打开失败",
                f"找不到这个文件：\n{path}\n\n"
                f"{reasons}\n\n"
                f"可以先点「选择文件」重新指定；\n"
                f"若反复出现，运行「{app_info.APP_NAME.replace(' ', '')}.exe --check」"
                f"生成环境自检报告。")
            return
        self.files = [path]
        name = os.path.basename(path)
        if self.current_mode == "detect":
            self._loaded_name = name
            self._draw_drop_card()
            self.detect_fonts()
        else:
            self._floaded_name = name
            self._draw_fdrop_card()
            self._set_fprog("文件已载入，点击「开始图片化」", name, 0)

    # ---------------- 字体检测 ----------------
    def choose_file(self):
        if self.detecting or self.installing:
            return
        p = self._ask_ppt_file("选择 PPT 文件")
        if p:
            self.files = [p]
            self._loaded_name = os.path.basename(p)
            self._draw_drop_card()
            self.detect_fonts()

    def _ask_ppt_file(self, title):
        """统一的打开对话框：initialdir 只给确实存在的目录，避免 Tk 报路径错。"""
        kw = {
            "title": title,
            "filetypes": [("PowerPoint 文件", "*.pptx *.pptm *.ppt *.pps *.ppsx *.potx"),
                          ("所有文件", "*.*")],
        }
        init = paths.safe_initialdir(
            os.path.dirname(self.files[0]) if self.files else self._fout_dir)
        if init:
            kw["initialdir"] = init
        p = filedialog.askopenfilename(**kw)
        return paths.normalize(p) if p else ""

    def detect_fonts(self, silent=False):
        if self.detecting or self.installing:
            return
        if not self.files:
            self.choose_file()
            return

        self.detecting = True
        self._silent = silent
        self._hint_text = None
        if not silent:
            self.btn_detect.set_text("检测中…")
            self.btn_detect.set_enabled(False)
        self._prog = 0.02
        self._draw_drop_card()
        for i in self.tree.get_children():
            self.tree.delete(i)

        threading.Thread(target=self._detect_worker,
                         args=(self.files[0], silent), daemon=True).start()

    def _detect_worker(self, path, silent):
        import time
        t0 = time.time()

        def progress(msg, cur=0, total=0):
            self._event_queue.put(("hint", msg))
            if total:
                self._event_queue.put(("prog", 0.05 + 0.9 * (cur / max(1, total))))

        try:
            result = font_detect.detect(path, progress_cb=progress,
                                        extra_font_names=self.tk_families)
            self._pending = (result, time.time() - t0)
            self._event_queue.put(("detect_done", None))
        except Exception as e:
            self._event_queue.put(("detect_fail", str(e)))

    def _detect_done(self):
        result, secs = self._pending
        self.fonts = result["fonts"]
        self.summary = result["summary"]
        self._detect_seconds = secs
        self._installed_total = self.summary.get("installed_total", 0)
        self._prog = None
        self._hint_text = None
        self.detecting = False
        self.btn_detect.set_text("检测字体")
        self.btn_detect.set_enabled(True)
        self._draw_env_chip()
        self._draw_drop_card(force=True)
        self._populate_tree()
        self._draw_list_card()
        self._draw_sum_card()

    def _populate_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)

        self.tree.tag_configure("installed", foreground="#9fb0c9")
        self.tree.tag_configure("missing", foreground="#ffa2b3")
        self.tree.tag_configure("sysmiss", foreground="#b9a273")
        self.tree.tag_configure("row_a", background=CARD_SOLID)
        self.tree.tag_configure("row_b", background="#161e2c")

        cat_map = {
            "opensource": ("开源免费", ACCENT),
            "free": ("免费商用", "#7fd0ff"),
            "commercial": ("商用授权", WARN),
            "system": ("系统自带", DIM),
            "unknown": ("未收录", DIM2),
        }
        for i, f in enumerate(self.fonts):
            label, _ = cat_map.get(f["category"], ("未知", DIM))
            info = f["info"]
            if isinstance(info, dict):
                if f["category"] == "opensource":
                    desc = f"{info.get('license', '')}"
                else:
                    desc = f"{info.get('vendor', '')} · {info.get('note', '')}"
            else:
                desc = str(info)

            if not f["installed"] and f["category"] == "opensource":
                desc += " · 可自动下载"
            elif not f["installed"] and f["category"] == "free":
                desc += " · 需手动下载"
            if f.get("embedded"):
                desc += " · 文档内嵌"

            matched = f.get("matched")
            if (f["installed"] and matched
                    and font_detect.norm_name(matched) != font_detect.norm_name(f["name"])):
                desc += f" · 本机命中「{matched}」"
            src = " / ".join(f.get("sources") or []) or "—"

            is_sys = f["category"] == "system"
            if f["installed"]:
                status, tag = "✓ 已安装", "installed"
            elif is_sys:
                status, tag = "· 系统", "sysmiss"
            else:
                status, tag = "✗ 缺失", "missing"
            self.tree.insert(
                "", "end",
                values=(f["name"], status, label, desc, src),
                tags=(tag, "row_b" if i % 2 else "row_a"))

    # ---------------- 悬停提示（复用同一个窗口，避免反复创建导致卡顿） ----------------
    def _ensure_tip(self):
        tip = getattr(self, "_tip", None)
        if tip is not None:
            try:
                if tip.winfo_exists():
                    return tip
            except Exception:
                pass
        tip = tk.Toplevel(self.root)
        tip.wm_overrideredirect(True)
        tip.configure(bg=BORDER_HL)
        lbl = tk.Label(tip, text="", bg="#1d2534", fg=TEXT, justify="left",
                       wraplength=self.px(420), font=self.fmt(9),
                       padx=self.px(10), pady=self.px(7))
        lbl.pack(padx=1, pady=1)
        tip._lbl = lbl
        tip.withdraw()
        self._tip = tip
        return tip

    def _tree_motion(self, event):
        row = self.tree.identify_row(event.y)
        if not row:
            self._hide_tip()
            return
        vals = self.tree.item(row, "values")
        if not vals:
            return
        text = "｜".join(str(v) for v in (vals[0], vals[2], vals[3]) if v)
        if text == getattr(self, "_tip_text", None):
            return
        self._tip_text = text
        tip = self._ensure_tip()
        tip._lbl.configure(text=text)
        tip.deiconify()
        tip.wm_geometry(f"+{event.x_root + self.px(14)}+{event.y_root + self.px(18)}")
        try:
            tip.lift()
        except Exception:
            pass

    def _hide_tip(self):
        self._tip_text = None
        tip = getattr(self, "_tip", None)
        if tip is not None:
            try:
                tip.withdraw()
            except Exception:
                self._tip = None

    def set_hint(self, msg):
        if not hasattr(self, "_hint_text"):
            self._hint_text = msg
        self._hint_text = msg

    def install_all(self):
        if self.installing or self.detecting:
            return
        todo = [f for f in self.fonts
                if not f["installed"] and f["category"] == "opensource"]
        if not todo:
            manual = [f for f in self.fonts
                      if not f["installed"] and f["category"] in ("free", "commercial", "unknown")]
            msg = "没有可自动下载的开源字体。"
            if manual:
                msg += "\n\n以下字体需手动获取（列表中有官方地址或厂商信息）：\n" + \
                       "、".join(f["name"] for f in manual[:8])
            messagebox.showinfo("提示", msg)
            return

        names = "、".join(f["name"] for f in todo)
        if not messagebox.askyesno(
                "确认下载安装",
                f"将自动下载并安装以下 {len(todo)} 款开源字体：\n{names}\n\n是否继续？"):
            return

        self.installing = True
        self.btn_install.set_text("安装中…")
        self.btn_install.set_enabled(False)
        self.btn_detect.set_enabled(False)
        threading.Thread(target=self._install_worker, args=(todo,), daemon=True).start()

    def _install_worker(self, todo):
        ok, fail = 0, 0
        for f in todo:
            key = (f["info"] or {}).get("key") if isinstance(f["info"], dict) else None
            key = key or f["name"]
            name = f["name"]

            def progress(m, n=name):
                self._event_queue.put(("hint", f"{n} · {m}"))

            try:
                font_install.download_and_install(key, progress_cb=progress)
                ok += 1
                f["installed"] = True
            except Exception as e:
                fail += 1
                f["_error"] = str(e)
                self._event_queue.put(("hint", f"{name} 失败：{e}"))

        self._event_queue.put(("install_done", (ok, fail)))

    def _install_done(self, ok, fail):
        self.installing = False
        self.btn_install.set_text("一键装缺字")
        self.btn_install.set_enabled(True)
        self.btn_detect.set_enabled(True)
        self.tk_families = list(tkfont.families())
        messagebox.showinfo("完成", f"成功安装 {ok} 款字体，失败 {fail} 款。\n\n已重新检测字体状态。")
        self.detect_fonts()

    def clear_all(self):
        if self.detecting or self.installing:
            return
        self.files = []
        self.fonts = []
        self.summary = {}
        self._loaded_name = None
        self._prog = None
        self._installed_total = 0
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._draw_env_chip()
        self._draw_drop_card()
        self._draw_list_card()
        self._draw_sum_card()

    # ---------------- 页面图片化 ----------------
    def fchoose_file(self):
        if self._flattening:
            return
        p = self._ask_ppt_file("选择 PPT 文件")
        if p:
            self.files = [p]
            self._floaded_name = os.path.basename(p)
            self._draw_fdrop_card()
            self._set_fprog("文件已载入，点击「开始图片化」", self._floaded_name, 0)

    def frun(self):
        if self._flattening:
            return
        if not self.files:
            self.fchoose_file()
            return

        src = self.files[0]
        dpi = self.dpi_var.get()
        out_dir = paths.safe_initialdir(self._fout_dir or os.path.dirname(src))
        base = os.path.splitext(os.path.basename(src))[0]
        # 文件名里的非法字符/首尾点空格会让保存对话框直接报路径错误
        base = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", base).strip().rstrip(".") or "输出"
        if len(base) > 80:
            base = base[:80]

        kw = {
            "title": "保存图片化后的 PPT（源文件不会被修改）",
            "defaultextension": ".pptx",
            "initialfile": base + "_图片化.pptx",
            "filetypes": [("PowerPoint 文件", "*.pptx")],
        }
        if out_dir:
            kw["initialdir"] = out_dir
        out = filedialog.asksaveasfilename(**kw)
        if not out:
            return
        if paths.same_file(out, src):
            messagebox.showerror("路径冲突", "输出文件不能与原文件相同，请换一个文件名。")
            return

        self._flattening = True
        self.btn_frun.set_text("转换中…")
        self.btn_frun.set_enabled(False)
        self.btn_fchoose.set_enabled(False)
        self._set_fprog("正在导出页面为高清图片…", f"{dpi} DPI · {os.path.basename(src)}", 0.02)
        threading.Thread(target=self._flatten_worker, args=(src, out, dpi),
                         daemon=True).start()

    def _flatten_worker(self, src, out, dpi):
        def progress(msg):
            ratio = None
            m = re.search(r"(?:导出|插入)第\s*(\d+)\s*/\s*(\d+)", msg)
            if m:
                cur, total = int(m.group(1)), max(1, int(m.group(2)))
                ratio = 0.05 + 0.9 * cur / total
            elif "生成" in msg or "保存" in msg:
                ratio = 0.96
            self._event_queue.put(("fprog", (msg, ratio)))

        try:
            slide_export.flatten_ppt(src, out, dpi=dpi, progress_cb=progress)
            self._event_queue.put(("flatten_done", out))
        except Exception as e:
            self._event_queue.put(("flatten_fail", str(e)))

    def _flatten_done(self, out):
        self._flattening = False
        self.btn_frun.set_text("开始图片化")
        self.btn_frun.set_enabled(True)
        self.btn_fchoose.set_enabled(True)
        self._set_fprog("图片化完成", os.path.basename(out), 1.0)
        messagebox.showinfo("完成",
                            f"页面图片化固定完成！\n\n输出文件：\n{out}\n\n"
                            f"每页已转为高清图片并插回新 PPT，\n"
                            f"任何设备打开都不会因缺字体而错位。")

    def _flatten_fail(self, err):
        self._flattening = False
        self.btn_frun.set_text("开始图片化")
        self.btn_frun.set_enabled(True)
        self.btn_fchoose.set_enabled(True)
        self._set_fprog("转换失败", err, 0)
        messagebox.showerror("转换失败", err)

    def fclear(self):
        if self._flattening:
            return
        self.files = []
        self._floaded_name = None
        self._draw_fdrop_card()
        self._set_fprog("等待开始…", "", 0)

    # ---------------- 事件循环（合并同一帧内的进度事件，避免重绘风暴） ----------------
    def _tick(self):
        hint = None
        prog = None
        try:
            while True:
                kind, payload = self._event_queue.get_nowait()
                if kind == "hint":
                    hint = payload
                elif kind == "prog":
                    prog = payload
                elif kind == "drop":
                    self._handle_dropped(payload)
                elif kind == "detect_done":
                    self._detect_done()
                elif kind == "detect_fail":
                    self.detecting = False
                    self._prog = None
                    self.btn_detect.set_text("检测字体")
                    self.btn_detect.set_enabled(True)
                    self._draw_drop_card()
                    messagebox.showerror("检测失败", payload)
                elif kind == "install_done":
                    self._install_done(*payload)
                elif kind == "fprog":
                    msg, ratio = payload
                    self._set_fprog(msg, self._fhint, ratio)
                elif kind == "flatten_done":
                    self._flatten_done(payload)
                elif kind == "flatten_fail":
                    self._flatten_fail(payload)
                elif kind == "env":
                    self._installed_total = payload
                    self._draw_env_chip()
        except queue.Empty:
            pass

        # 仅在仍在检测时消费进度（否则会把上一轮的残留进度画回来）
        if self.detecting and (prog is not None or hint is not None):
            if prog is not None:
                self._prog = prog
            if hint is not None:
                self.set_hint(hint)
            self._draw_drop_card()

        self.root.after(60, self._tick)


class SelfTestDialog:
    """功能自检窗口：跑内置样本，逐项给出结论。

    引擎在后台线程里跑（自检会真的打开一次 Office 导图，可能十几秒），
    主线程只负责把队列里的进度画出来，避免窗口假死。
    """

    MARKS = {
        "ok": ("● 通过", "ok"),
        "warn": ("● 注意", "warn"),
        "fail": ("● 失败", "fail"),
        "skip": ("● 跳过", "skip"),
    }

    def __init__(self, app):
        self.app = app
        self.win = None
        self.queue = queue.Queue()
        self.report = None
        self.running = False

    # ------------------------------------------------------------ 构建
    def open(self):
        if self.win is not None:
            try:
                if self.win.winfo_exists():
                    self.win.deiconify()
                    self.win.lift()
                    self.win.focus_force()
                    return
            except tk.TclError:
                pass
            self.win = None
        self._build()
        self.start()

    def _build(self):
        app, px, fmt = self.app, self.app.px, self.app.fmt
        wa = work_area()
        w, h = px(790), px(700)
        if wa[2] > 400:
            w = min(w, wa[2] - px(40))
            h = min(h, wa[3] - px(50))

        win = tk.Toplevel(app.root)
        win.title("功能自检 · 内置测试 PPT")
        win.configure(bg=BG)
        win.transient(app.root)
        win.protocol("WM_DELETE_WINDOW", self.close)
        win.bind("<Escape>", lambda e: self.close())

        app.root.update_idletasks()
        rx, ry = app.root.winfo_rootx(), app.root.winfo_rooty()
        rw, rh = max(400, app.root.winfo_width()), max(400, app.root.winfo_height())
        x = rx + (rw - w) // 2
        y = ry + (rh - h) // 2
        if wa[2] > 400:
            x = max(wa[0] + px(8), min(x, wa[0] + wa[2] - w - px(8)))
            y = max(wa[1] + px(8), min(y, wa[1] + wa[3] - h - px(8)))
        win.geometry("%dx%d+%d+%d" % (w, h, x, y))
        win.minsize(px(620), px(460))
        self.win = win

        head = tk.Frame(win, bg=BG)
        head.pack(fill="x", padx=px(24), pady=(px(18), px(10)))
        tk.Label(head, text="功能自检", bg=BG, fg=TEXT,
                 font=fmt(16, "bold")).pack(anchor="w")
        tk.Label(head,
                 text="用内置的 3 份测试 PPT 真实跑一遍：资源释放 → 路径 → 字体检测 → 自动保存 → 图片化",
                 bg=BG, fg=DIM, font=fmt(9)).pack(anchor="w", pady=(px(5), 0))

        body = tk.Frame(win, bg=BG)
        body.pack(fill="both", expand=True, padx=px(24))
        self.txt = tk.Text(
            body, bg="#0d141f", fg=TEXT, relief="flat", bd=0,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=BORDER,
            padx=px(14), pady=px(10), wrap="word", font=fmt(9),
            insertwidth=0, cursor="arrow", spacing1=px(1), spacing3=px(1))
        sb = tk.Scrollbar(body, command=self.txt.yview, width=px(11), bd=0,
                          relief="flat", bg="#1b2333", activebackground=BORDER_HL,
                          troughcolor="#0d141f", highlightthickness=0)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)
        self.txt.configure(yscrollcommand=sb.set)

        t = self.txt
        t.tag_configure("ok", foreground=OK, font=fmt(9, "bold"))
        t.tag_configure("fail", foreground=FAIL, font=fmt(9, "bold"))
        t.tag_configure("warn", foreground=WARN, font=fmt(9, "bold"))
        t.tag_configure("skip", foreground=DIM2, font=fmt(9, "bold"))
        t.tag_configure("title", foreground=TEXT, font=fmt(10, "bold"),
                        spacing1=px(9))
        t.tag_configure("det", foreground=DIM, lmargin1=px(20), lmargin2=px(20),
                        font=fmt(9))
        t.tag_configure("hint", foreground=DIM2, font=fmt(8), lmargin1=px(20),
                        lmargin2=px(20), spacing3=px(4))
        t.configure(state="disabled")

        foot = tk.Frame(win, bg=BG)
        foot.pack(fill="x", padx=px(24), pady=(px(12), px(16)))
        self.lbl = tk.Label(foot, text="准备中…", bg=BG, fg=DIM, font=fmt(9),
                            anchor="w", justify="left")
        self.lbl.pack(side="left", fill="x", expand=True)

        btns = tk.Frame(foot, bg=BG)
        btns.pack(side="right")
        self.btn_again = app._pill(btns, "重新检测", self.start, "ghost", 100, 36, 9)
        self.btn_again.pack(side="left", padx=px(4))
        self.btn_files = app._pill(btns, "测试文件", self.open_files, "ghost", 96, 36, 9)
        self.btn_files.pack(side="left", padx=px(4))
        self.btn_report = app._pill(btns, "打开报告", self.open_report, "primary", 104, 36, 9)
        self.btn_report.pack(side="left", padx=px(4))
        close = app._pill(btns, "关闭", self.close, "ghost", 72, 36, 9)
        close.pack(side="left", padx=(px(4), 0))

    # ------------------------------------------------------------ 执行
    def start(self):
        if not (self.win and self.win.winfo_exists()):
            return
        if self.running:
            return
        self.running = True
        self.report = None
        t = self.txt
        t.configure(state="normal")
        t.delete("1.0", "end")
        t.configure(state="disabled")
        self.lbl.configure(text="正在检测…", fg=DIM)
        self.btn_again.set_enabled(False)
        self.btn_files.set_enabled(False)
        self.btn_report.set_enabled(False)
        threading.Thread(target=self._worker, daemon=True).start()
        self._poll()

    def _worker(self):
        try:
            import selftest
            selftest.run(cb=lambda kind, payload: self.queue.put((kind, payload)))
        except Exception:
            self.queue.put(("crash", traceback.format_exc()))

    def _poll(self):
        if not (self.win and self.win.winfo_exists()):
            return
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                self._handle(kind, payload)
        except queue.Empty:
            pass
        if self.win is not None:
            try:
                self.win.after(70, self._poll)
            except tk.TclError:
                pass

    def _handle(self, kind, payload):
        if kind == "phase":
            self.lbl.configure(text=str(payload), fg=DIM)
        elif kind == "step":
            self._append_step(payload)
        elif kind == "done":
            self._finish(payload)
        elif kind == "crash":
            self._write("自检过程中断：\n", "fail")
            self._write(str(payload).strip() + "\n", "det")
            self.running = False
            self.lbl.configure(text="自检异常中断", fg=FAIL)
            self._enable_after()

    def _write(self, text, tag=None):
        t = self.txt
        t.configure(state="normal")
        t.insert("end", text, (tag,) if tag else ())
        t.configure(state="disabled")
        t.see("end")

    def _append_step(self, step):
        mark, tag = self.MARKS.get(step.get("status"), ("●", "det"))
        self._write(mark + "  ", tag)
        self._write(step.get("title", "") + "\n", "title")
        for line in step.get("lines", []):
            self._write("    " + line + "\n", "det")

    def _enable_after(self):
        try:
            self.btn_again.set_enabled(True)
            self.btn_files.set_enabled(bool(self.report and self.report.get("work_dir")))
            self.btn_report.set_enabled(bool(self.report and self.report.get("report_path")))
        except Exception:
            pass

    def _finish(self, report):
        self.report = report
        self.running = False
        steps = report.get("steps", [])
        fails = [s for s in steps if s.get("status") == "fail"]
        skips = [s for s in steps if s.get("status") == "skip"]
        ok = sum(1 for s in steps if s.get("status") == "ok")

        self._write("\n" + "─" * 34 + "\n", "hint")
        if fails:
            self._write("发现 %d 项失败，对应功能可能不可用\n" % len(fails), "fail")
            for s in fails:
                self._write("    · %s\n" % s.get("title", ""), "det")
        elif skips:
            self._write("%d 项通过，%d 项跳过（环境不具备，不影响主要功能）\n"
                        % (ok, len(skips)), "warn")
        else:
            self._write("全部通过（%d 项），功能链路正常\n" % ok, "ok")
        if report.get("report_path"):
            self._write("报告：" + report["report_path"] + "\n", "hint")
        if report.get("work_dir"):
            self._write("测试文件：" + report["work_dir"] + "\n", "hint")

        if fails:
            self.lbl.configure(text="检测完成：%d 项失败 · 耗时 %s"
                               % (len(fails), report.get("elapsed", "-")), fg=FAIL)
        else:
            self.lbl.configure(text="检测完成：%d 项通过 · 耗时 %s"
                               % (ok, report.get("elapsed", "-")), fg=OK)
        self._enable_after()

    # ------------------------------------------------------------ 动作
    def open_report(self):
        path = (self.report or {}).get("report_path")
        if not path:
            return
        try:
            paths.open_path(path)
        except Exception as exc:
            messagebox.showerror("打开报告",
                                 "打不开报告文件：\n%s\n\n%s" % (path, exc),
                                 parent=self.win)

    def open_files(self):
        path = (self.report or {}).get("work_dir")
        if not path:
            return
        try:
            paths.open_path(path)
        except Exception as exc:
            messagebox.showerror("打开目录",
                                 "打不开测试文件目录：\n%s\n\n%s" % (path, exc),
                                 parent=self.win)

    def close(self):
        try:
            if self.win is not None:
                self.win.destroy()
        except Exception:
            pass
        self.win = None


def _install_error_handler(root=None):
    """未捕获异常不再弹英文堆栈框：写日志 + 给一句人话 + 指向功能自检。"""

    def handle(exc_type, exc, tb):
        detail = "".join(traceback.format_exception(exc_type, exc, tb))
        target = ""
        try:
            target = paths.log_file("crash.log")
            with paths.open_binary(target, "a", encoding="utf-8") as f:
                f.write("[%s] v%s frozen=%s\n%s\n%s\n"
                        % (time.strftime("%Y-%m-%d %H:%M:%S"), app_info.VERSION,
                           paths.is_frozen(), detail, "-" * 66))
        except Exception:
            target = ""
        if os.environ.get("PPTDEBUG") == "1":
            sys.__excepthook__(exc_type, exc, tb)
            return
        try:
            messagebox.showerror(
                "程序遇到一个错误",
                "刚才的操作没有完成：\n%s: %s\n\n" % (exc_type.__name__, exc)
                + ("详细信息已写入：\n%s\n\n" % target if target else "")
                + "可以到「关于」页点「一键功能自检」，用内置测试 PPT 检查环境。",
                parent=root if root is not None else None)
        except Exception:
            pass

    sys.excepthook = handle
    if root is not None:
        # Tk 回调里的异常默认会被 Tk 自己吞掉并打印到 stderr
        # （窗口程序没有 stderr，就会变成 PyInstaller 的崩溃框）
        root.report_callback_exception = handle


def main():
    # 旧快捷方式的"起始位置"可能指向已删除的盘/目录（换电脑后极常见），
    # 这会让 os.getcwd() 直接抛 WinError 3，任何相对路径都用不了。
    paths.ensure_cwd()
    enable_dpi_awareness()
    try:
        log_path = paths.log_file("startup.log")
        if log_path:
            with open(paths.long_path(log_path), "a", encoding="utf-8") as f:
                f.write("[%s] v%s frozen=%s exe=%s cwd=%s\n" % (
                    time.strftime("%Y-%m-%d %H:%M:%S"), app_info.VERSION,
                    paths.is_frozen(), paths.exe_path(), os.getcwd()))
    except Exception:
        pass
    root = tk.Tk()
    _install_error_handler(root)
    if not MACOS:
        # macOS 的 Tk 坐标本来就是逻辑点（72dpi），改 scaling 会把字号放大 33%，
        # 自绘界面里文字就会超出卡片，所以只在 Windows 上设。
        try:
            dpi = max(float(root.winfo_fpixels("1i")), float(system_dpi()))
            root.tk.call("tk", "scaling", dpi / 72.0)
        except Exception:
            pass
    app = PPTToolbox(root)
    # 换了目录或换了电脑后，注册表里的自启动路径可能已失效 —— 就地修正
    try:
        if autostart.sync():
            app.as_engine.log("info", "开机启动项路径已更新为当前程序位置")
    except Exception:
        pass
    if "--autostart" in sys.argv[1:]:
        root.after(200, app.enter_background_mode)
    root.mainloop()


def _cli_flatten(argv):
    """隐藏自检入口（窗口化 EXE 没有 stdout，因此把过程写进日志文件）。

    用法：PPT工具箱.exe --flatten <源文件> [输出文件] [dpi]
    """
    import tempfile
    import traceback

    log_path = paths.log_file("pptfx_cli.log")
    if not log_path:
        log_path = os.path.join(tempfile.gettempdir(), "pptfx_cli.log")

    def log(msg):
        try:
            with open(paths.long_path(log_path), "a", encoding="utf-8") as f:
                f.write("[%s] %s\n" % (time.strftime("%H:%M:%S"), msg))
        except Exception:
            pass

    src = paths.resolve_existing(argv[0]) if argv else ""
    out = argv[1] if len(argv) > 1 else os.path.splitext(src)[0] + "_图片化.pptx"
    try:
        dpi = int(argv[2]) if len(argv) > 2 else 192
    except Exception:
        dpi = 192
    try:
        open(paths.long_path(log_path), "w").close()
        log("src=%s out=%s dpi=%d" % (src, out, dpi))
        log("src_exists=%s" % paths.is_file(src))
        log("cwd=%s" % os.getcwd())
        log("powerpoint=%r" % slide_export.find_powerpoint_exe())
        slide_export.flatten_ppt(src, out, dpi=dpi, progress_cb=log)
        log("OK -> %s (%d bytes)" % (out, os.path.getsize(out)))
        return 0
    except Exception:
        log("FAIL\n" + traceback.format_exc())
        return 1


def _self_check(gui=None, open_file=True):
    """环境自检：换电脑后先跑这个，报告里能直接看出是哪一环的路径有问题。"""
    if gui is None:
        # 打包成窗口程序后没有控制台，只能弹窗；源码运行时有 stdout 就打印
        gui = paths.is_frozen() or sys.stdout is None
    try:
        target, text = paths.write_diagnostic_report()
    except Exception as exc:
        target, text = "", "自检失败：%r" % (exc,)
    if gui:
        try:
            messagebox.showinfo(
                "环境自检",
                text + ("\n报告已保存：\n%s" % target if target else ""))
        except Exception:
            pass
    else:
        try:
            sys.stdout.write(text)
            if target:
                sys.stdout.write("\n报告：%s\n" % target)
        except Exception:
            pass
    if open_file and target:
        try:
            paths.open_path(target)
        except Exception:
            pass
    return target


# CI 只认「打包出来就该有」的项。Office / LibreOffice 缺失在构建机上很正常，
# 所以不拿 diagnose() 的全量结果当闸门，只挑这几项。
_CI_MUST = ("程序位置", "资源 logo.png", "资源 logo_mark.png",
            "用户数据目录", "临时目录", "字体目录", "字体索引缓存")


def _cli_ci():
    """给 GitHub Actions 用的产物探针。

    和 --check 的区别：--check 是给人看的全量报告，缺 Office 也照常退出 0；
    这里只验打包产物自身是否完整，缺关键项就返回 1 让 CI 失败。
    文本同时写入 $CI_REPORT（打包成窗口程序后 stdout 是 None，CI 里读文件更稳）。
    """
    lines = ["PPT 工具箱 · 产物探针",
             "版本 %s · %s" % (app_info.VERSION, app_info.PLATFORM),
             "打包运行：%s" % ("是" if paths.is_frozen() else "否（源码）"),
             ""]
    fail = 0
    try:
        rows = paths.diagnose()
    except Exception as exc:
        rows = []
        lines.append("[FAIL] 诊断本身出错：%r" % (exc,))
        fail += 1
    for name, ok, detail in rows:
        if name not in _CI_MUST:
            continue
        lines.append("%s %s：%s" % ("[ OK ]" if ok else "[FAIL]", name, detail))
        if not ok:
            fail += 1

    # 内嵌自检样本必须活着：PyInstaller 漏了 hiddenimports 时这里会先响
    try:
        import selftest_data
        n = len(selftest_data.SELFTEST_FILES)
        lines.append("%s 内置自检样本：%d 份" % ("[ OK ]" if n >= 3 else "[FAIL]", n))
        fail += 0 if n >= 3 else 1
    except Exception as exc:
        lines.append("[FAIL] 内置自检样本：%r" % (exc,))
        fail += 1

    # 只建 Tcl 解释器，不建主窗口（CI 上没有可用的窗口会话）
    try:
        import tkinter
        tkinter.Tcl()
        lines.append("[ OK ] tkinter：Tk %s" % tkinter.TkVersion)
    except Exception as exc:
        lines.append("[FAIL] tkinter：%r" % (exc,))
        fail += 1

    lines += ["", "结论：%s" % ("通过" if fail == 0 else "%d 项未通过" % fail)]
    text = "\n".join(lines) + "\n"

    try:
        if sys.stdout is not None:
            sys.stdout.write(text)
            sys.stdout.flush()
    except Exception:
        pass
    target = os.environ.get("CI_REPORT", "")
    if target:
        try:
            with open(target, "w", encoding="utf-8") as fh:
                fh.write(text)
        except Exception:
            pass
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    # 打包成窗口程序后 stdout/stderr 是 None，任何 print / write 都会抛
    # AttributeError 并弹出“Unhandled exception”，所以命令行分支统一走这里。
    def _say(text):
        try:
            if sys.stdout is not None:
                sys.stdout.write(text)
                sys.stdout.flush()
        except Exception:
            pass

    if len(sys.argv) > 1 and sys.argv[1] == "--flatten":
        sys.exit(_cli_flatten(sys.argv[2:]))
    if len(sys.argv) > 1 and sys.argv[1] == "--ci":
        sys.exit(_cli_ci())
    if len(sys.argv) > 1 and sys.argv[1] in ("--check", "--selfcheck"):
        _self_check(open_file=False)
        sys.exit(0)
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        import selftest
        res = selftest.run(lambda kind, data: None)
        _say(res["text"] + "\n")
        sys.exit(0 if res["fail"] == 0 else 1)
    main()
