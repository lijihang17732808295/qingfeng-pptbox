# -*- coding: utf-8 -*-
"""
PPT 字体检测核心模块（极速版）

字体来源（多路交叉，互不遗漏，全程不依赖 Office）：
  1. OOXML 深度扫描：所有 slide/layout/master/theme/chart/diagram/notes 的
     latin/ea/cs/sym typeface + 主题字体引用(+mj-lt / +mn-ea …) 还原
  2. 文档内嵌字体：<p:embeddedFont> 与 ppt/fonts/*.fntdata
  3. .ppt 二进制：FontEntityAtom(0x0FC8) 记录解析 + 常见字体名清单兜底

性能要点：
  · XML 一律按字节流正则匹配，命中的 typeface 才解码，避免整份 UTF-8 解码
  · 字体文件只 seek 读 name 表（几 KB），不再整文件读入（省掉数百 MB IO）
  · 归一化 / 分类 / 匹配全部预编译 + 记忆化，索引在进程内单例复用
"""
import os
import re
import sys
import json
import struct
import html
import bisect
import threading
import unicodedata
import zipfile
from collections import defaultdict
from functools import lru_cache

import font_db
import paths

# 兼容旧接口
OPEN_SOURCE_FONTS = font_db.OPEN_SOURCE_FONTS

if sys.platform == "win32":
    import winreg

FONT_EXTS = (".ttf", ".ttc", ".otf", ".otc")

THEME_KEYS = {
    "+mj-lt": "mj-lt", "+mj-ea": "mj-ea", "+mj-cs": "mj-cs", "+mj-sym": "mj-sym",
    "+mn-lt": "mn-lt", "+mn-ea": "mn-ea", "+mn-cs": "mn-cs", "+mn-sym": "mn-sym",
}

# 字重/风格词：用于判断"前缀相同但只是字重不同"的字体是否算已装
STYLE_WORDS = frozenset({
    "regular", "normal", "book", "roman", "std", "standard", "ui", "text", "display",
    "light", "ultralight", "extralight", "semilight", "demilight", "thin", "hairline",
    "bold", "semibold", "demibold", "medium", "heavy", "black", "extrabold", "ultrabold",
    "italic", "oblique", "condensed", "expanded", "narrow", "wide", "caption", "small",
    "cn", "sc", "tc", "hk", "gb", "gb2312", "gbk", "big5", "pro", "h", "hw", "mono",
    "variable", "vf", "w3", "w4", "w5", "w6", "w7", "w8", "w9", "l", "m", "b", "sb",
    "jp", "kr", "tw", "mo", "alt", "alternate", "new",
})
_STYLE_SORTED = tuple(sorted(STYLE_WORDS, key=len, reverse=True))


# ============================================================
# 0. 名称规范化 / 别名
# ============================================================
def _pc(cb, msg, cur=0, total=0):
    """兼容两种进度回调签名"""
    if not cb:
        return
    try:
        cb(msg, cur, total)
    except TypeError:
        try:
            cb(msg)
        except Exception:
            pass
    except Exception:
        pass


_NON_NAME_RE = re.compile(r"[^0-9a-z\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff"
                          r"\uac00-\ud7af\u1100-\u11ff\uf900-\ufaff]")


@lru_cache(maxsize=200000)
def norm_name(s: str) -> str:
    """归一化：全角转半角、小写、去掉一切标点空格（结果缓存，重复调用零成本）"""
    if not s:
        return ""
    return _NON_NAME_RE.sub("", unicodedata.normalize("NFKC", s).lower())


def _strip_weight(s: str) -> str:
    return s


@lru_cache(maxsize=None)
def _only_style_words(rest: str) -> bool:
    """判断剩余片段是否只由字重/变体词组成（arialblack -> arial + black）"""
    if not rest:
        return True
    if rest.isdigit():
        return True
    i, n = 0, len(rest)
    got = False
    while i < n:
        for w in _STYLE_SORTED:
            if rest.startswith(w, i):
                i += len(w)
                got = True
                break
        else:
            return False
    return got


def build_alias_map() -> dict:
    """归一化名 -> 同组所有归一化名（含自身）"""
    m = defaultdict(set)
    for group in font_db.ALIAS_GROUPS:
        norms = [norm_name(g) for g in group if g]
        norms = [n for n in norms if n]
        for n in norms:
            m[n].update(norms)
    return m


ALIAS_MAP = build_alias_map()


# ============================================================
# 1. 字体文件 name 表解析（只读必需片段）
# ============================================================
_NAME_IDS = frozenset((1, 4, 16, 21))  # Family / Full / Typographic Family / WWS Family
_MAX_NAME_TABLE = 8 << 20             # name 表上限 8MB，防御异常文件


def _decode_name(raw: bytes, plat: int) -> str:
    try:
        if plat in (0, 3):      # Unicode / Windows: UTF-16BE
            return raw.decode("utf-16-be", "ignore")
        if plat == 1:           # Mac Roman
            return raw.decode("mac_roman", "ignore")
    except Exception:
        pass
    return ""


def _parse_name_table(table: bytes) -> set:
    names = set()
    if len(table) < 6:
        return names
    try:
        count, str_off = struct.unpack(">HH", table[2:6])
    except Exception:
        return names
    for i in range(min(count, 512)):
        r = 6 + i * 12
        if r + 12 > len(table):
            break
        try:
            plat, enc, lang, nid, length, off = struct.unpack(">HHHHHH", table[r:r + 12])
        except Exception:
            continue
        if nid not in _NAME_IDS or not length:
            continue
        s = _decode_name(table[str_off + off:str_off + off + length], plat).strip()
        if s and len(s) < 96 and "?" not in s:
            names.add(s)
    return names


def _names_at(f, offset: int) -> set:
    """在文件 f 的 offset 处读取一个 sfnt 的 name 表"""
    names = set()
    try:
        f.seek(offset)
        hdr = f.read(12)
        if len(hdr) < 12:
            return names
        num = struct.unpack(">H", hdr[4:6])[0]
        if not (0 < num < 512):
            return names
        dir_data = f.read(16 * num)
        name_off = name_len = None
        for i in range(min(num, len(dir_data) // 16)):
            rec = dir_data[i * 16:(i + 1) * 16]
            if rec[:4] == b"name":
                name_off, name_len = struct.unpack(">II", rec[8:16])
                break
        if name_off is None or not (0 < name_len <= _MAX_NAME_TABLE):
            return names
        f.seek(name_off)
        return _parse_name_table(f.read(name_len))
    except Exception:
        return names


def read_font_names(path: str) -> set:
    """读取字体文件内所有家族名（支持 TTC 字体集合，只读 name 表片段）"""
    try:
        with paths.open_binary(path, "rb") as f:
            head = f.read(12)
            if len(head) < 12:
                return set()
            if head[:4] != b"ttcf":
                return _names_at(f, 0)
            f.seek(8)
            num = struct.unpack(">I", f.read(4))[0]
            f.seek(12)
            # 先把所有子字体偏移全部读出（后续 _names_at 会移动文件指针）
            raw_offs = f.read(4 * min(num, 64))
            names = set()
            for i in range(0, len(raw_offs) - 3, 4):
                off = struct.unpack(">I", raw_offs[i:i + 4])[0]
                names |= _names_at(f, off)
            return names
    except Exception:
        return set()


# ============================================================
# 2. 本机已装字体索引
# ============================================================
class FontIndex:
    def __init__(self):
        self.raw = set()      # 原始字体名（展示用）
        self.norm = {}        # 归一化 -> 原始名
        self.loaded_from_cache = False
        self._sorted = None
        self._memo = {}

    def add(self, name):
        if not name:
            return
        name = name.strip()
        if not name or len(name) > 96:
            return
        n = norm_name(name)
        if not n:
            return
        self.raw.add(name)
        if n not in self.norm:
            self.norm[n] = name
            self._sorted = None
            self._memo = {}

    def __len__(self):
        return len(self.raw)

    # ---------- 匹配 ----------
    def _keys(self):
        if self._sorted is None:
            self._sorted = sorted(self.norm)
        return self._sorted

    def match(self, font_name: str):
        """返回本机匹配到的字体名，未装返回 None（结果记忆化）"""
        cached = self._memo.get(font_name, 0)
        if cached != 0:
            return cached
        res = self._match(font_name)
        if len(self._memo) > 20000:
            self._memo.clear()
        self._memo[font_name] = res
        return res

    def _match(self, font_name: str):
        cands = [norm_name(font_name)]
        for c in list(cands):
            if c in ALIAS_MAP:
                cands.extend(ALIAS_MAP[c])
        cands = [c for c in dict.fromkeys(cands) if c]

        # 1. 精确
        for c in cands:
            raw = self.norm.get(c)
            if raw is not None:
                return raw
        if not self.norm:
            return None
        keys = self._keys()

        for c in cands:
            # 2. 本机名以候选开头，剩余只是字重/变体（二分定位前缀区间）
            i = bisect.bisect_left(keys, c)
            while i < len(keys) and keys[i].startswith(c):
                k = keys[i]
                if len(k) > len(c) and _only_style_words(k[len(c):]):
                    return self.norm[k]
                i += 1
            # 3. 候选以本机名开头，剩余只是字重/变体（按长度枚举前缀，不做全表扫描）
            for L in range(3, len(c)):
                inst = c[:L]
                raw = self.norm.get(inst)
                if raw is not None and _only_style_words(c[L:]):
                    return raw
            # 4. 长名互相包含（>=5，避免 heiti/黑体 之类误判）：枚举候选子串查表
            n = len(c)
            for L in range(n, 4, -1):
                for s in range(0, n - L + 1):
                    raw = self.norm.get(c[s:s + L])
                    if raw is not None:
                        return raw
        return None


def _registry_names() -> list:
    out = []
    if sys.platform != "win32":
        return out
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        try:
            with winreg.OpenKey(hive, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts") as k:
                i = 0
                while True:
                    try:
                        name, value, _ = winreg.EnumValue(k, i)
                    except OSError:
                        break
                    i += 1
                    clean = re.sub(r"\s*\((TrueType|OpenType|All res)\)\s*$", "", name).strip()
                    # "Microsoft YaHei & Microsoft YaHei UI" 这类合并名要拆开
                    for part in re.split(r"\s*&\s*|\s*、\s*", clean):
                        part = part.strip()
                        if part:
                            out.append(part)
                    if value:
                        stem = os.path.splitext(os.path.basename(value))[0]
                        if stem:
                            out.append(stem)
        except OSError:
            pass
    return out


def _font_files() -> list:
    out = []
    for d in paths.font_dirs():
        try:
            for fn in os.listdir(paths.long_path(d)):
                if fn.lower().endswith(FONT_EXTS):
                    out.append(os.path.join(d, fn))
        except OSError:
            pass
    return out


def _cache_path() -> str:
    """索引缓存文件；拿不到可写目录时返回空串（调用方跳过缓存）。"""
    try:
        return paths.cache_file("font_index.json") or ""
    except Exception:
        return ""


_INDEX_LOCK = threading.Lock()
_INDEX_MEMO = {"key": None, "idx": None}


def _index_key(files) -> str:
    total = 0
    for p in files:
        try:
            total += os.path.getsize(p)
        except OSError:
            pass
    # 把字体目录也纳入 key：换机器/换用户后即使数量与体积碰巧相同，
    # 目录不同也不会误用上一台机器的缓存。
    try:
        import zlib
        dirs = "\n".join(sorted({os.path.dirname(p).lower() for p in files}))
        stamp = "%08x" % (zlib.crc32(dirs.encode("utf-8", "ignore")) & 0xFFFFFFFF)
    except Exception:
        stamp = "0"
    return "%d|%d|%s" % (len(files), total, stamp)


def get_installed_fonts(progress_cb=None, extra_names=None, use_cache=True) -> FontIndex:
    """
    建立本机字体索引（进程内单例：同一份字体目录只解析一次）。
    progress_cb(msg, cur, total) 可选
    extra_names: 额外字体名（如 tkinter.font.families() 的结果）
    """
    files = _font_files()
    key = _index_key(files)

    with _INDEX_LOCK:
        memo = _INDEX_MEMO
        if use_cache and memo["key"] == key and memo["idx"] is not None:
            idx = memo["idx"]
        else:
            idx = FontIndex()
            if use_cache:
                cache_path = _cache_path()
                if cache_path:
                    try:
                        with paths.open_binary(cache_path, "r",
                                               encoding="utf-8") as f:
                            cache = json.load(f)
                        if cache.get("key") == key:
                            for n in cache.get("names", []):
                                idx.add(n)
                            idx.loaded_from_cache = True
                    except Exception:
                        pass

            if not idx.loaded_from_cache:
                _pc(progress_cb, "读取系统字体清单…", 0, 1)
                for n in _registry_names():
                    idx.add(n)

                total = max(1, len(files))
                for i, p in enumerate(files):
                    if i % 25 == 0:
                        _pc(progress_cb, f"解析字体文件 {i + 1}/{total}", i, total)
                    idx.add(os.path.splitext(os.path.basename(p))[0])
                    for nm in read_font_names(p):
                        idx.add(nm)
                cache_path = _cache_path()
                if cache_path:
                    try:
                        with paths.open_binary(cache_path, "w",
                                               encoding="utf-8") as f:
                            json.dump({"key": key, "names": sorted(idx.raw)}, f,
                                      ensure_ascii=False)
                    except Exception:
                        pass

            memo["key"], memo["idx"] = key, idx
            if progress_cb:
                _pc(progress_cb, f"本机字体 {len(idx)} 款", 1, 1)

    for n in (extra_names or []):
        idx.add(n)
    return idx


# ============================================================
# 3. 从 PPT 中提取字体（全部基于字节流，命中才解码）
# ============================================================
_RE_TYPEFACE = re.compile(rb'typeface="([^"]{1,200})"')
_RE_EMBED_BLOCK = re.compile(rb"<p:embeddedFont[^>]*>(.*?)</p:embeddedFont>", re.S)
_RE_EMBED_FONT = re.compile(rb'<p:font[^>]*typeface="([^"]+)"')
_RE_THEME = {}
for _tag in ("majorFont", "minorFont"):
    _t = _tag.encode("ascii")
    _RE_THEME[_tag] = re.compile(b"<a:" + _t + b">(.*?)</a:" + _t + b">", re.S)
_RE_THEME_FONT = re.compile(rb'<a:font\s+script="[^"]*"\s+typeface="([^"]+)"')


def _b2s(raw: bytes) -> str:
    return html.unescape(raw.decode("utf-8", "ignore")).strip()


def _parse_theme(raw: bytes) -> dict:
    """从 theme xml（字节）取出 major/minor 主题字体"""
    res = {}
    for kind, tag in (("mj", "majorFont"), ("mn", "minorFont")):
        m = _RE_THEME[tag].search(raw)
        if not m:
            continue
        body = m.group(1)
        for sub, key in ((b"latin", "lt"), (b"ea", "ea"), (b"cs", "cs"), (b"sym", "sym")):
            mm = re.search(rb'<a:%s[^>]*typeface="([^"]*)"' % sub, body)
            if mm and mm.group(1).strip():
                res["%s-%s" % (kind, key)] = _b2s(mm.group(1))
        for mm in _RE_THEME_FONT.finditer(body):
            v = _b2s(mm.group(1))
            if v:
                res.setdefault("%s-ea" % kind, v)
    return res


def scan_pptx(path, add, progress_cb=None):
    """深度扫描 OOXML（.pptx/.pptm/.ppsx/.potx）"""
    with zipfile.ZipFile(paths.long_path(path)) as z:
        parts = z.namelist()

        # 主题字体（可能多套）
        union = defaultdict(set)
        for n in parts:
            if "theme" in n and n.endswith(".xml"):
                try:
                    t = _parse_theme(z.read(n))
                except Exception:
                    continue
                for k, v in t.items():
                    union[k].add(v)

        xmls = [n for n in parts if n.lower().endswith(".xml")]
        total = max(1, len(xmls))
        for i, n in enumerate(xmls):
            if i % 20 == 0:
                _pc(progress_cb, f"扫描页面结构 {i + 1}/{total}", i, total)
            try:
                raw = z.read(n)
            except Exception:
                continue
            if b"typeface" not in raw and b"embeddedFont" not in raw:
                continue

            # 内嵌字体声明
            if b"embeddedFont" in raw:
                for m in _RE_EMBED_BLOCK.finditer(raw):
                    for f in _RE_EMBED_FONT.findall(m.group(1)):
                        add(_b2s(f), "内嵌")

            for m in _RE_TYPEFACE.finditer(raw):
                nm = _b2s(m.group(1))
                if not nm:
                    continue
                if nm.startswith("+"):
                    for v in union.get(THEME_KEYS.get(nm, ""), ()):
                        add(v, "主题")
                    continue
                add(nm, "XML")

        # 有 fntdata 说明确实内嵌了字体文件
        if any(n.lower().endswith(".fntdata") for n in parts):
            add("__EMBEDDED_FONT_FILES__", "内嵌文件")


def _valid_font_name(s: str) -> bool:
    if not s or len(s) > 40:
        return False
    if any(ch in s for ch in "，。！？；：、“”（）《》\t\r\n"):
        return False
    if "  " in s or s.isdigit():
        return False
    if any("\u3400" <= ch <= "\u9fff" or "\u3040" <= ch <= "\u30ff" for ch in s):
        return True
    if not s[0].isupper():
        return False
    return all(ch.isalnum() or ch in " -._" for ch in s)


def scan_ppt_binary(path, add, progress_cb=None):
    """扫描 .ppt 二进制：FontEntityAtom 记录 + 已知字体名清单"""
    try:
        with paths.open_binary(path, "rb") as f:
            data = f.read()
    except Exception:
        return

    hits = 0
    try:
        pos = 0
        find = data.find
        while True:
            pos = find(b"\xc8\x0f", pos)
            if pos < 0:
                break
            rec_len = int.from_bytes(data[pos + 2:pos + 6], "little")
            if 0x20 <= rec_len <= 0x80:
                name = data[pos + 6:pos + 6 + 64].decode("utf-16-le", "ignore")
                name = name.split("\x00")[0].strip()
                if _valid_font_name(name):
                    add(name, "二进制")
                    hits += 1
            pos += 2
        _pc(progress_cb, f"解析字体记录 {hits} 条", 1, 1)
    except Exception:
        pass

    # 已知字体名清单扫描（仅当记录解析不足时兜底，避免误报）
    if hits >= 2:
        return
    for fam in font_db.KNOWN_FONT_NAMES:
        try:
            if fam.encode("utf-16-le") in data:
                add(fam, "二进制")
            elif fam.isascii() and len(fam) >= 5 and fam.encode("ascii") in data:
                add(fam, "二进制")
        except Exception:
            continue


def extract_fonts(path, progress_cb=None) -> tuple:
    """
    返回 (fonts:list[str], sources:dict[name]->set(tag), notes:list[str])
    """
    path = paths.resolve_existing(path)
    if not paths.is_file(path):
        reasons = "；".join(paths.describe_missing(path))
        raise FileNotFoundError(f"文件不存在：{path}\n{reasons}")

    ext = os.path.splitext(path)[1].lower()
    if ext not in (".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx"):
        raise ValueError("不是有效的 PowerPoint 文件")

    # 可读性检查（文件可能正被 PowerPoint 独占打开）
    try:
        with paths.open_binary(path, "rb") as f:
            head = f.read(8)
    except Exception:
        raise RuntimeError("无法读取该文件，可能正被 PowerPoint 或其他程序占用，请先关闭后重试")
    if not head:
        raise RuntimeError("文件为空或读取失败")

    sources = defaultdict(set)
    order = []

    def add(name, tag):
        name = (name or "").strip()
        if name == "__EMBEDDED_FONT_FILES__":
            sources["__EMBEDDED__"].add("内嵌")
            return
        if not name or len(name) > 80 or name.startswith("+"):
            return
        if name not in sources:
            order.append(name)
        sources[name].add(tag)

    notes = []

    if ext in (".pptx", ".pptm", ".ppsx", ".potx"):
        try:
            _pc(progress_cb, "深度扫描文档结构…", 0, 1)
            scan_pptx(path, add, progress_cb)
        except Exception as e:
            notes.append(f"XML 扫描失败：{e}")
    else:
        try:
            _pc(progress_cb, "扫描二进制字体记录…", 0, 1)
            scan_ppt_binary(path, add, progress_cb)
        except Exception as e:
            notes.append(f"二进制扫描失败：{e}")

    if not order:
        notes.append("未在文件中解析到字体信息，可能为加密或损坏的文件")

    return order, sources, notes


# 兼容旧接口
def extract_ppt_fonts(ppt_path, progress_cb=None) -> list:
    fonts, _, _ = extract_fonts(ppt_path, progress_cb)
    return sorted(fonts, key=lambda s: s.lower())


# ============================================================
# 4. 字体分类（关键词表预归一化，避免每次重算）
# ============================================================
def _build_classify_tables():
    open_tab = []
    for key, val in font_db.OPEN_SOURCE_FONTS.items():
        kws = [k for k in (norm_name(x) for x in val["keywords"]) if k]
        open_tab.append((kws, key, val))

    free_tab = []
    for item in font_db.FREE_FONTS:
        kws = [k for k in (norm_name(x) for x in item["keywords"]) if k]
        free_tab.append((kws, item))

    sys_restr = [(norm_name(k), d) for k, d in font_db.SYSTEM_RESTRICTED.items()
                 if norm_name(k)]
    sys_latin = sorted((norm_name(k) for k in font_db.SYSTEM_LATIN if norm_name(k)),
                       key=len, reverse=True)
    brands = []
    for kws, vendor, note in font_db.COMMERCIAL_BRANDS:
        nk = [k for k in (norm_name(x) for x in kws) if k]
        brands.append((nk, vendor, note))
    known = [k for k in (norm_name(x) for x in font_db.KNOWN_FONT_NAMES) if k]
    return open_tab, free_tab, sys_restr, sys_latin, brands, known


(_OPEN_TAB, _FREE_TAB, _SYS_RESTR, _SYS_LATIN,
 _BRANDS, _KNOWN_NORM) = _build_classify_tables()


@lru_cache(maxsize=100000)
def classify_font(font_name: str):
    """返回 (category, info)
    category: opensource | free | commercial | system | unknown
    """
    name = (font_name or "").strip()
    n = norm_name(name)
    if not n:
        return ("unknown", "字体名为空")

    # 1. 开源（可自动下载）
    for kws, key, val in _OPEN_TAB:
        for k in kws:
            if k in n:
                return ("opensource", {
                    "family": val["family"], "license": val["license"],
                    "url": val["url"], "repo": val.get("repo"), "key": key,
                })

    # 2. 免费商用（需手动下载）
    for kws, item in _FREE_TAB:
        for k in kws:
            if k in n:
                return ("free", {
                    "name": item["name"], "vendor": item["vendor"],
                    "url": item["url"], "note": item["note"],
                })

    # 3. 系统中文（商用受限）
    for k, desc in _SYS_RESTR:
        if n == k or (n.startswith(k) and _only_style_words(n[len(k):])):
            return ("commercial", desc)

    # 4. 系统/Office 自带西文
    for k in _SYS_LATIN:
        if n == k or (n.startswith(k) and _only_style_words(n[len(k):])):
            return ("system", "系统 / Office 自带，一般无需处理")

    # 5. 商业字库（短关键词只做前缀匹配，避免 "AR" 命中 Barlow 之类误判）
    for kws, vendor, note in _BRANDS:
        for k in kws:
            if n == k or n.startswith(k) or (len(k) >= 3 and k in n):
                return ("commercial", f"{vendor} · {note}")

    # 6. 未识别
    return ("unknown", "未收录字体，请自行确认授权")


# ============================================================
# 5. 综合检测入口
# ============================================================
def detect(ppt_path, progress_cb=None, extra_font_names=None) -> dict:
    import time

    t0 = time.time()
    _pc(progress_cb, "解析文档字体…", 0, 1)
    names, sources, notes = extract_fonts(ppt_path, progress_cb)
    t_parse = time.time() - t0

    t0 = time.time()
    _pc(progress_cb, "读取本机字体…", 0, 1)
    idx = get_installed_fonts(progress_cb=progress_cb, extra_names=extra_font_names)
    t_index = time.time() - t0

    t0 = time.time()
    _pc(progress_cb, "比对与分类…", 0, 1)
    result = []
    for name in names:
        matched = idx.match(name)
        cat, info = classify_font(name)
        tags = sources.get(name, set())
        result.append({
            "name": name,
            "installed": matched is not None,
            "matched": matched,
            "category": cat,
            "info": info,
            "sources": sorted(tags),
            "embedded": "内嵌" in tags,
        })
    t_match = time.time() - t0

    # 缺失优先，其次按名称
    result.sort(key=lambda r: (r["installed"], r["category"] == "system",
                               r["name"].lower()))

    missing = [r for r in result if not r["installed"]]
    summary = {
        "total": len(result),
        "installed_count": len(result) - len(missing),
        "missing_count": len(missing),
        "opensource_missing": [r for r in missing if r["category"] == "opensource"],
        "free_missing": [r for r in missing if r["category"] == "free"],
        "commercial_missing": [r for r in missing if r["category"] == "commercial"],
        "unknown_missing": [r for r in missing if r["category"] == "unknown"],
        "system_missing": [r for r in missing if r["category"] == "system"],
        "system_missing_count": sum(1 for r in missing if r["category"] == "system"),
        "embedded_count": sum(1 for r in result if r["embedded"]),
        "installed_total": len(idx),
        "notes": notes,
        "timing": {"parse": t_parse, "index": t_index, "match": t_match},
    }
    return {"fonts": result, "summary": summary}


if __name__ == "__main__":
    import sys as _s
    if len(_s.argv) > 1:
        r = detect(_s.argv[1], progress_cb=lambda m, *a: print("  ·", m))
        print("\n=== 检测结果 ===")
        for f in r["fonts"]:
            mark = "已装" if f["installed"] else "缺失"
            print(f"[{mark}] {f['name']:<28} {f['category']:<11} 来源={'/'.join(f['sources'])}")
        s = r["summary"]
        print(f"\n共 {s['total']} 种，已装 {s['installed_count']}，缺失 {s['missing_count']}")
        print(f"本机字体总数 {s['installed_total']}，内嵌 {s['embedded_count']}")
        for n in s["notes"]:
            print("注：", n)
