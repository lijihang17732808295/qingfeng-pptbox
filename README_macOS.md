# PPT 工具箱 · macOS 版（v1.07）

同一份代码，Windows 打包成 `.exe`，macOS 打包成 `.app`。界面、字体检测、页面图片化的用法完全一致。

> **重要前提**：PyInstaller 不支持交叉编译。`.app` 只能在 macOS 上生成，
> 在 Windows 上无论怎么打包都得不到能在 Mac 上运行的二进制文件。
> 所以这里交付的是**已完成 macOS 移植的源码 + 一键打包脚本**，在任意一台 Mac 上执行一条命令即可产出 `.app`。

---

## 一、在 Mac 上打包（约 3 分钟）

```bash
# 1. 把整个目录拷到 Mac，进入目录
cd /path/to/ppt-font-checker

# 2. 执行打包脚本
bash build_mac.sh
```

脚本会自动完成：建虚拟环境 → 装依赖 → 生成 `app.icns` → PyInstaller 打包 → ad-hoc 签名 → 压缩。

产物：

```
dist/PPT工具箱.app          ← 双击即可运行，可拖进「应用程序」
dist/PPT工具箱-mac.zip      ← 拷给别人用这个
```

**环境要求**

| 项目 | 要求 |
|---|---|
| 系统 | macOS 10.15 (Catalina) 及以上，Intel / Apple Silicon 均可 |
| Python | 3.9–3.12，且带 tkinter（推荐 python.org 官方安装包；Homebrew 需 `brew install python-tk`） |
| Xcode 命令行工具 | `xcode-select --install`（PyInstaller 需要） |

**不打包，直接用源码跑**（改代码调试时更快）：

```bash
python3 -m pip install -r requirements-mac.txt
python3 main.py
```

---

## 二、功能对照

| 功能 | Windows | macOS | 说明 |
|---|---|---|---|
| 字体检测（版式/母版/主题/图表/内嵌字体） | ✅ | ✅ | 纯 OOXML 解析，不依赖 Office；mac 索引 `/System/Library/Fonts`、`/Library/Fonts`、`~/Library/Fonts` |
| 缺失字体一键下载 | ✅ | ✅ | GitHub Release 下载 |
| 字体安装 | ✅ 用户字体目录 + 注册表 | ✅ 复制到 `~/Library/Fonts` | mac 不需要管理员密码 |
| 页面图片化 | ✅ PowerPoint COM | ✅ LibreOffice → PDF → PNG，或 PowerPoint 自动化 | 输出仍为「每页一张整页图」的 pptx |
| 自动保存 | ✅ PowerPoint + WPS | ⚠️ 仅 Microsoft PowerPoint | 走 AppleScript，首次会请求「自动化」授权；WPS for Mac 无脚本接口 |
| 定时备份 | ✅ SaveCopyAs | ✅ 复制磁盘上的文件 | 备份的是已保存到磁盘的版本 |
| 开机自启动 | ✅ 注册表 Run | ✅ LaunchAgent | 不带管理员权限 |
| 拖拽文件 | ✅ 原生 WM_DROPFILES | ❌ Tk 收不到系统拖放 | 改用「选择文件 / 载入文件」按钮 |
| 功能自检（内置 3 份测试 PPT） | ✅ | ✅ | 图片化那一项在 mac 上会先找 LibreOffice，其次 PowerPoint |

### macOS 版的行为差异

1. **页面图片化需要一个导出引擎**：优先用 LibreOffice（免费、无需授权），没装则用 Microsoft PowerPoint。
   两者都没有时，程序会明确告诉你缺什么，字体检测不受影响。
   LibreOffice 下载：<https://zh-cn.libreoffice.org/>
2. **首次用 PowerPoint 自动化**：系统会弹「"PPT工具箱"想要控制"Microsoft PowerPoint"」，必须点**好**，否则自动保存/导出会失败。
   误点了「不允许」→ 系统设置 → 隐私与安全性 → 自动化 → 重新勾选。
3. **备份目录**：默认 `~/Documents/PPT自动备份`，可在「自动保存」页改。
4. **配置与日志**：`~/Library/Application Support/PPTToolbox/`
   （`autosave.json`、`font_index.json`、`crash.log`、`selftest_*.txt`）
5. **不支持加密/受保护的 PPT**，与 Windows 版一致。

---

## 三、命令行参数（排错用）

```bash
# 环境自检，会生成一份报告
"dist/PPT工具箱.app/Contents/MacOS/PPT工具箱" --check

# 功能自检（6 项：资源/路径/字体索引/字体检测/自动保存/图片化）
"dist/PPT工具箱.app/Contents/MacOS/PPT工具箱" --selftest

# 直接跑一次图片化
"dist/PPT工具箱.app/Contents/MacOS/PPT工具箱" --flatten 源文件.pptx 输出.pptx 192
```

`PPTDEBUG=1` 会额外把过程写进日志，遇到问题先看 `~/Library/Application Support/PPTToolbox/`。

---

## 四、首次打开被拦怎么办

未做 Apple 开发者签名（需要 99 美元/年的证书），系统会提示"无法验证开发者"：

- 方式一：在「访达」里**右键 → 打开 → 打开**
- 方式二：`xattr -dr com.apple.quarantine "/Applications/PPT工具箱.app"`
- 方式三：系统设置 → 隐私与安全性 → 仍要打开

如果要把 App 分发给别人，建议自行做 Developer ID 签名 + 公证：

```bash
codesign --force --deep --options runtime --sign "Developer ID Application: 你的名字 (TEAMID)" "dist/PPT工具箱.app"
xcrun notarytool submit "dist/PPT工具箱-mac.zip" --apple-id ... --team-id ... --password ... --wait
xcrun stapler staple "dist/PPT工具箱.app"
```

---

## 五、这次改了哪些文件

| 文件 | 改动 |
|---|---|
| `paths.py` | macOS 目录规则：Application Support、字体目录、Documents、`open_path()` 跨平台打开；诊断项按平台分叉 |
| `font_install.py` | mac 安装到 `~/Library/Fonts` + `atsutil` 刷新 |
| `autostart.py` | mac 用 LaunchAgent plist（`launchctl bootstrap/bootout`），路径变化自动重写 |
| `autosave.py` | mac 走 AppleScript 后端（列文档 / 保存 / 复制备份），WPS 明确不支持 |
| `slide_export.py` | mac 导出链：LibreOffice→PDF→PyMuPDF 渲染；PowerPoint AppleScript 兜底；python-pptx 拼页 |
| `main.py` | 字体回退（PingFang SC）、DPI 与 tk scaling 分支、拖放按平台降级、文案与按钮按平台调整 |
| `selftest.py` | 报告系统字段、图片化引擎识别按平台分叉 |
| `app_info.py` | 版本 1.07，界面显示「macOS 版 / Windows 版」 |
| `PPT工具箱_mac.spec`、`build_mac.sh`、`requirements-mac.txt` | 新增，macOS 打包用 |

Windows 版行为不变，仍用原来的 `PPT工具箱.spec` 打包。

---

## 六、已知风险

- 本机没有 Mac，**这份移植没有在真机上跑过**。纯 Python 部分（路径层、字体检测、python-pptx 拼页）在 Windows 上做了等价验证；
  macOS 专有的三块（AppleScript 自动保存、LaunchAgent 自启动、osascript 导出）属于按官方脚本字典写的实现，**首次在 Mac 上使用时请先点「关于 → 一键功能自检」**，
  报告里会逐项说明哪一环没通。
- 若自检里「图片化导出」失败，先装 LibreOffice 再试；这条路比 PowerPoint 自动化稳定得多。
