# -*- coding: utf-8 -*-
"""
应用元信息（单点维护，界面与打包共用）
"""
import sys

APP_NAME = "PPT 工具箱"
APP_TITLE = "PPT 工具箱 · 字体检测 & 图片化 & 自动保存"
VERSION = "1.07"

# 平台标识：Windows 打包成 exe，macOS 打包成 .app，界面代码共用一份
PLATFORM = ("macOS" if sys.platform == "darwin"
            else ("Windows" if sys.platform == "win32" else sys.platform))
PLATFORM_TAG = "macOS 版" if sys.platform == "darwin" else "Windows 版"

DEVELOPER = "清风BHU"
TAGLINE = "公益工具"

SUBTITLE = "字体检测 · 页面图片化 · 自动保存" + (
    " · 拖拽即用" if sys.platform == "win32" else "")

ABOUT_LINES = [
    ("开发者", DEVELOPER),
    ("版本", "V%s · %s" % (VERSION, PLATFORM_TAG)),
    ("性质", TAGLINE),
    ("说明", "免费使用，禁止倒卖与二次收费；转载须保留作者与版本信息"),
]
