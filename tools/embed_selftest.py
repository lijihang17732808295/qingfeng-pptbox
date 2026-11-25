# -*- coding: utf-8 -*-
"""把 assets/selftest/*.pptx 压成 selftest_data.py（开发期工具，不打包）。"""
import base64
import glob
import gzip
import os
import textwrap
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "assets", "selftest")
DST = os.path.join(ROOT, "selftest_data.py")

META = {
    "01_基础字体与文本.pptx": (
        "直接写在 run 上的字体，含中英日韩字符",
        ("微软雅黑", "宋体", "黑体", "楷体", "仿宋")),
    "02_主题母版与表格.pptx": (
        "主题字体、母版占位符、表格、组合形状、图表",
        ("等线", "楷体")),
    "03_特殊场景与边界.pptx": (
        "旋转、越界、超长文本、罕见字体与特殊符号",
        ("华文行楷", "楷体", "黑体")),
}


def main():
    parts = [
        "# -*- coding: utf-8 -*-",
        '"""',
        "内置自检用 PPT（gzip + base64），随 EXE 一起打包。",
        "",
        "为什么要把文件塞进代码里：用户换电脑后最常见的故障是"
        "“路径不存在 / 资源丢失”，",
        "此时再去让用户找一个能用的 PPT 来试，等于跳过问题本身。"
        "内置样本可以保证",
        "自检永远有输入，且内容固定、期望值已知，结果可对比。",
        "",
        "由 tools/embed_selftest.py 生成，请勿手改。",
        '"""',
        "",
        "SELFTEST_FILES = [",
    ]
    total = 0
    for path in sorted(glob.glob(os.path.join(SRC, "*.pptx"))):
        name = os.path.basename(path)
        raw = open(path, "rb").read()
        total += len(raw)
        packed = gzip.compress(raw, 9)
        b64 = base64.b64encode(packed).decode("ascii")
        slides = len([n for n in zipfile.ZipFile(path).namelist()
                      if n.startswith("ppt/slides/slide")
                      and n.endswith(".xml")])
        desc, expect = META[name]
        parts.append("    {")
        parts.append('        "name": %r,' % name)
        parts.append('        "desc": %r,' % desc)
        parts.append('        "size": %d,' % len(raw))
        parts.append('        "slides": %d,' % slides)
        parts.append('        "expect": %r,' % (list(expect),))
        parts.append('        "b64": (')
        for line in textwrap.wrap(b64, 100):
            parts.append('            "%s"' % line)
        parts.append("        ),")
        parts.append("    },")
        print("%-26s raw=%6d gzip=%6d b64=%6d slides=%d"
              % (name, len(raw), len(packed), len(b64), slides))
    parts.append("]")
    parts.append("")
    with open(DST, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(parts))
    print("total raw = %d bytes; %s = %d bytes"
          % (total, os.path.basename(DST), os.path.getsize(DST)))


if __name__ == "__main__":
    main()
