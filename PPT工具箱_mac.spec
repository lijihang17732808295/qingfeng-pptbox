# -*- mode: python ; coding: utf-8 -*-
"""
macOS 打包配置（PyInstaller）

必须在 macOS 上执行，不能从 Windows 交叉编译：
    pyinstaller --clean --noconfirm PPT工具箱_mac.spec

产物：dist/PPT工具箱.app
"""

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[('assets/logo.png', 'assets'),
           ('assets/logo_mark.png', 'assets')],
    # 这些模块是按需 import 的，静态分析抓不到，必须显式列出来
    hiddenimports=[
        'anchor_data', 'selftest', 'selftest_data',
        'autosave', 'autostart', 'paths', 'font_db', 'font_detect',
        'font_install', 'slide_export',
        'pptx', 'pptx.util', 'pptx.parts', 'pptx.oxml',
        'lxml', 'lxml.etree', 'PIL._tkinter_finder',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['win32com', 'pythoncom', 'pywintypes', 'winreg', '_winreg'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PPT工具箱',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='PPT工具箱',
)

app = BUNDLE(
    coll,
    name='PPT工具箱.app',
    icon='app.icns',
    bundle_identifier='com.qingfengbh.ppttoolbox',
    version='1.07',
    info_plist={
        'CFBundleName': 'PPT工具箱',
        'CFBundleDisplayName': 'PPT 工具箱',
        'CFBundleShortVersionString': '1.07',
        'CFBundleVersion': '1.07',
        'LSMinimumSystemVersion': '10.15',
        'NSHighResolutionCapable': True,
        'NSRequiresAquaSystemAppearance': False,
        # 自动保存 / 图片化 需要驱动 PowerPoint，系统会弹「自动化」授权框
        'NSAppleEventsUsageDescription':
            '需要控制 Microsoft PowerPoint 来完成自动保存与页面图片化导出。',
        'NSDesktopFolderUsageDescription': '读取你选择的 PPT 文件并写入导出结果。',
        'NSDocumentsFolderUsageDescription': '读取你选择的 PPT 文件并写入导出结果。',
        'NSDownloadsFolderUsageDescription': '读取你选择的 PPT 文件并写入导出结果。',
    },
)
