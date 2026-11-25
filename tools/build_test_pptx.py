# -*- coding: utf-8 -*-
"""
生成内置自检用的测试 PPT（只在开发期跑，不参与打包）。

产物：assets/selftest/*.pptx
用途：换电脑 / 路径异常时，用这几份固定文件验证
      「资源释放 → 字体检测 → 图片化导出」整条链路是否正常。

三份文件各自覆盖不同的解析路径：
  01 基础字体与文本   —— 直接写在 run 上的字体、项目符号、多语言字符
  02 主题母版与表格   —— 主题字体、母版/版式、表格、备注、组合形状
  03 特殊场景与边界   —— 旋转、超出页面、透明、超长文本、罕见字体
"""
import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "assets", "selftest")

INK = RGBColor(0x1A, 0x1A, 0x1A)
GREY = RGBColor(0x6B, 0x72, 0x80)
ACCENT = RGBColor(0x2A, 0x9D, 0x8F)


def _blank(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _text(slide, x, y, w, h, lines, font="宋体", size=20, bold=False,
          color=INK, align=PP_ALIGN.LEFT, spacing=1.2):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        r = p.add_run()
        r.text = line
        r.font.name = font
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
    return box


def build_basic():
    """01：run 级字体 + 项目符号 + 中英日韩字符"""
    prs = Presentation()
    s = _blank(prs)
    _text(s, 0.9, 0.7, 8.2, 1.0, ["字体检测 · 基础用例"], "微软雅黑", 34, True, INK)
    _text(s, 0.9, 1.7, 8.2, 0.5, ["用于验证 run 级字体能否被完整读出"], "微软雅黑", 14, False, GREY)

    _text(s, 0.9, 2.5, 3.6, 2.6, [
        "· 宋体正在这里出现",
        "· 这是第二行正文",
        "· 数字 0123456789",
        "· 标点，。；：「」",
    ], "宋体", 18)

    _text(s, 4.9, 2.5, 3.6, 2.6, [
        "黑体加粗标题样",
        "黑体常规正文样",
        "Mixed 中英混排 Text",
        "— 破折号与省略号……",
    ], "黑体", 18)

    _text(s, 0.9, 5.3, 8.2, 1.2, [
        "Latin words keep their own metrics: Leadership, Copyright, Evidence.",
        "日本語のテキストとハングル 텍스트도 함께 넣었습니다.",
    ], "等线", 13, False, GREY)

    s = _blank(prs)
    _text(s, 0.9, 0.7, 8.2, 0.8, ["第二页 · 段落与对齐"], "微软雅黑", 26, True)
    _text(s, 0.9, 1.6, 8.2, 1.6, [
        "这一页用楷体，验证同一文件内多种字体能否分别归类。",
        "第二段继续使用楷体，行距与首行缩进不影响字体识别。",
    ], "楷体", 17, spacing=1.6)
    _text(s, 0.9, 3.4, 8.2, 1.2, ["仿宋在这段文字里出现，仅作为字体样本。"],
          "仿宋", 17)
    _text(s, 0.9, 4.6, 8.2, 1.6, [
        "右对齐样本：",
        "字体检测只读，不会修改文件内容。",
    ], "宋体", 15, align=PP_ALIGN.RIGHT)
    return prs


def build_theme():
    """02：主题字体 + 母版版式 + 表格 + 备注 + 组合形状 + 图表"""
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[0])
    s.shapes.title.text = "主题与母版用例"
    s.placeholders[1].text = "标题页占位符会继承主题字体"

    s = _blank(prs)
    _text(s, 0.7, 0.5, 8.6, 0.7, ["表格与备注"], "等线", 24, True)

    data = [
        ("项目", "字体", "说明"),
        ("第一行", "宋体", "表格内文字也要能解析"),
        ("第二行", "等线", "等线是 Office 默认中文字体"),
        ("第三行", "Consolas", "等宽字体用于代码片段"),
    ]
    table = s.shapes.add_table(len(data), 3, Inches(0.7), Inches(1.3),
                               Inches(8.6), Inches(2.3)).table
    for r, row in enumerate(data):
        for cidx, val in enumerate(row):
            cell = table.cell(r, cidx)
            cell.text = val
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(13)
                    run.font.name = "等线" if cidx == 1 else "宋体"

    # 组合形状：组合内部的文字单独指定字体
    group = s.shapes.add_group_shape()
    rect = group.shapes.add_shape(1, Inches(0.7), Inches(4.0),
                                  Inches(3.6), Inches(1.0))
    rect.text_frame.text = "组合形状里的文字"
    for p in rect.text_frame.paragraphs:
        for run in p.runs:
            run.font.name = "楷体"
            run.font.size = Pt(16)
            run.font.bold = True

    s.notes_slide.notes_text_frame.text = "备注里也写一种字体：华文中宋，用于验证备注解析。"

    s = _blank(prs)
    _text(s, 0.7, 0.6, 8.6, 0.8, ["第三页 · 图表"], "等线", 24, True)
    cd = CategoryChartData()
    cd.categories = ["首页", "内页", "尾页"]
    cd.add_series("页数", (1, 3, 1))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                            Inches(0.7), Inches(1.5), Inches(4.8), Inches(3.3), cd)
    chart = gf.chart
    chart.has_legend = False
    chart.chart_title.text_frame.text = "图表标题也会带字体"
    return prs


def build_special():
    """03：旋转、越界、透明、超长文本、罕见字体"""
    prs = Presentation()
    s = _blank(prs)

    _text(s, 0.8, 0.6, 8.4, 0.8, ["特殊场景与边界"], "华文行楷", 32, True)

    # 旋转文本框
    box = _text(s, 1.0, 1.8, 3.2, 1.6, ["旋转了 12 度的文本框", "字体为楷体"],
                "楷体", 16)
    box.rotation = 12

    # 超出页面左边界的元素（模拟版式溢出）
    over = _text(s, -0.6, 3.6, 3.0, 1.0, ["这段文字有一半在页面外"], "宋体", 15)
    over.left = Emu(-548640)

    # 透明形状 + 等线文字
    shape = s.shapes.add_shape(1, Inches(4.8), Inches(1.8), Inches(4.2), Inches(1.8))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0xF2, 0xF6, 0xF8)
    shape.line.color.rgb = RGBColor(0xC9, 0xD3, 0xDC)
    tf = shape.text_frame
    tf.word_wrap = True
    tf.text = "形状内的等线文字，含全角空格与「直角引号」"
    for p in tf.paragraphs:
        for run in p.runs:
            run.font.name = "等线"
            run.font.size = Pt(15)

    # 超长无空格文本（测试换行与截断）
    long_line = "超长连续文本用于验证解析不会截断" * 4
    _text(s, 0.9, 5.4, 8.4, 1.2, [long_line], "黑体", 12)

    s = _blank(prs)
    _text(s, 0.8, 0.6, 8.4, 0.8, ["字符集覆盖"], "微软雅黑", 26, True)
    _text(s, 0.8, 1.6, 8.4, 4.6, [
        "中文标点：，。、；：？！“”‘’（）《》【】—…·「」『』",
        "符号与数学：±×÷≈≠≤≥∞∑√∫°′″",
        "货币与单位：￥＄€￡‰℃㎏",
        "全角与半角混排：ＡＢＣ１２３ abc123",
        "表情与箭头：→←↑↓★☆●○■□◆◇",
        "生僻字：龘靐齉爩麤龗灪厵",
    ], "宋体", 15, spacing=1.5)
    return prs


def main():
    os.makedirs(OUT, exist_ok=True)
    items = [
        ("01_基础字体与文本.pptx", build_basic),
        ("02_主题母版与表格.pptx", build_theme),
        ("03_特殊场景与边界.pptx", build_special),
    ]
    for name, builder in items:
        path = os.path.join(OUT, name)
        prs = builder()
        prs.save(path)
        print("%-26s %7d bytes  %d slides" %
              (name, os.path.getsize(path), len(prs.slides._sldIdLst)))


if __name__ == "__main__":
    main()
