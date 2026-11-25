# PPT 工具箱

> 字体检测 · 页面图片化 · 定时自动保存 —— 一个不用装 Office 也能查 PPT 字体的桌面小工具。
> Windows 打包成 `.exe`，macOS 打包成 `.app`，界面与功能共用同一份代码。

[![Build macOS App](https://github.com/lijihang17732808295/qingfeng-pptbox/actions/workflows/build-macos.yml/badge.svg)](https://github.com/lijihang17732808295/qingfeng-pptbox/actions/workflows/build-macos.yml)
![Version](https://img.shields.io/badge/version-1.07-blue)
![Python](https://img.shields.io/badge/python-3.9%20%7C%203.10%20%7C%203.11%20%7C%203.12-3776ab?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-lightgrey)

---

## 它解决什么问题

拿到一份别人做的 PPT，打开发现字体全变了、版式全乱了，却不知道原稿用了哪些字体、哪个是母版带来的、哪个只出现在图表的坐标轴里。PowerPoint 自己的「替换字体」对话框只告诉你顶层字体，母版、版式、主题、SmartArt、图表、内嵌字体的引用它一概看不到。

这个小工具直接解析 OOXML（`.pptx` 本质是 zip），把**每一处**字体引用都翻出来，并对照本机字体库标出缺哪些。全程不依赖 Office，也不用联网。

另外顺手做了两件周边的事：把 PPT 每页导成图片重新拼一份（换机器不跑版），以及盯住正在编辑的演示文稿按间隔自动保存 + 另存备份。

## 功能

| 功能 | 说明 | Windows | macOS |
|---|---|:--:|:--:|
| **字体检测** | 扫描版式 / 母版 / 主题 / 图表 / SmartArt / 内嵌字体，标出本机缺失项，给出替换建议 | ✅ | ✅ |
| **缺失字体下载** | 从 GitHub Release 拉取开源替代字体（思源、Noto 等），可批量 | ✅ | ✅ |
| **字体安装** | 装到当前用户字体目录，不需要管理员权限 | ✅ | ✅ |
| **页面图片化** | 每页渲染成一张整页图，重新拼成 pptx；原稿字体缺失也不会跑版 | ✅ | ✅ |
| **定时自动保存** | 按间隔对正在编辑的文档执行「保存」，等同 Ctrl+S | ✅ | ⚠️ 仅 PowerPoint |
| **定时备份** | 按间隔另存独立副本，不改动原文件，可设保留份数 | ✅ | ✅ |
| **开机自启动** | 登录后静默在后台监听，可随时开关 | ✅ | ✅ |
| **一键自检** | 内置 3 份测试 PPT，报错时点一下就知道是哪一环坏了 | ✅ | ✅ |
| **拖拽文件** | 把 PPT 拖进窗口即开始解析 | ✅ | ❌ |

> macOS 上「自动保存」走 AppleScript 驱动 Microsoft PowerPoint，首次会请求系统「自动化」授权；WPS for Mac 没有脚本接口，因此不支持。拖拽在 macOS 上不可用（Tk 收不到系统拖放事件），改用「选择文件」按钮。

## 平台差异的根因

不是套了个壳，关键路径是真的分叉了：

| 环节 | Windows | macOS |
|---|---|---|
| 数据目录 | `%LOCALAPPDATA%\PPTToolbox`（Known Folder，兼容 OneDrive 重定向） | `~/Library/Application Support` |
| 字体目录 | `C:\Windows\Fonts` + 用户字体目录 | `/System/Library/Fonts`、`/Library/Fonts`、`~/Library/Fonts` |
| 字体安装 | 复制 + 写注册表 | 复制到 `~/Library/Fonts` + `atsutil fonts -update` |
| 页面图片化 | PowerPoint COM 导出 PNG | LibreOffice → PDF → PyMuPDF 渲染；没装则退到 PowerPoint 自动化 |
| 自动保存 | COM 附着 `PowerPoint.Application` / `KWPP.Application` | osascript 脚本字典 |
| 开机自启 | `HKCU\...\CurrentVersion\Run` | `~/Library/LaunchAgents/*.plist` + `launchctl` |
| 超长路径 | `\\?\` 前缀 + 短路径名兜底 | 无 260 限制，仍需处理 NFD 归一化 |

## 下载使用

**macOS（Apple Silicon，M 系列）**：直接下这个，不用登录、不用编译

```
https://github.com/lijihang17732808295/qingfeng-pptbox/releases/latest/download/PPTToolbox-v1.07-macOS-arm64.zip
```

解压后把「PPT工具箱.app」拖进「应用程序」即可。

- 首次打开若提示「无法验证开发者」→ 右键 → 打开（只需一次）
- 或执行：`xattr -dr com.apple.quarantine /Applications/PPT工具箱.app`
- 首次用「页面图片化」或「自动保存」会弹「想要控制 Microsoft PowerPoint」，点**好**

**Windows**：`PPT字体检测器.exe`，双击即用（单文件，无需安装 Python）

> 2019 年及更早的 **Intel MacBook** 不能跑 arm64 包。这种情况到 Actions 页手动 Run workflow
> 勾一次 x86_64 构建，跑完用 `gh release upload v1.07 dist/PPT工具箱-mac.zip` 补上传。

## 从源码运行

```bash
# Windows
python -m venv .venv && .venv\Scripts\activate
pip install -r requirements.txt
python main.py

# macOS
python3 -m venv .venv-mac && source .venv-mac/bin/activate
pip install -r requirements-mac.txt
python3 main.py
```

## 自己打包

### Windows

```bash
pyinstaller --clean --noconfirm PPT工具箱.spec
# → dist/PPT字体检测器.exe
```

### macOS

```bash
bash build_mac.sh
# → dist/PPT工具箱.app、dist/PPT工具箱-mac.zip
```

脚本会自动建虚拟环境、装依赖、用 `sips`/`iconutil` 生成 `.icns`、PyInstaller 打包、ad-hoc 签名、压缩。要求 macOS 10.15+、带 tkinter 的 Python（推荐 python.org 官方包）、`xcode-select --install`。

> PyInstaller 不能交叉编译：`.app` 只能在 macOS 上生成，`.exe` 只能在 Windows 上生成。这也是下面这条 CI 存在的理由。

## 自动构建（GitHub Actions）

`.github/workflows/build-macos.yml` 在 GitHub 提供的 macOS 机器上把 `.app` 打出来，本地没有 Mac 也能拿到安装包。

**触发条件**

| 时机 | 行为 |
|---|---|
| 推送到 `main` | 语法检查 + 打包 arm64，产物存为 artifact（保留 30 天） |
| 推送 `v*` 标签 | 创建 Release，挂上 arm64 zip |
| Pull Request | 只验证能不能编译通过 |
| 手动 Run workflow | 双架构重跑（不发版） |

**job 的分工**

1. `lint` —— 在 `ubuntu-latest` 上先 `compileall`，语法错就不用排队等 macOS 机器。
2. `arm64` —— `macos-15`（Apple Silicon），每次 push 都跑。
3. `intel` —— `macos-13`（Intel），只在手动触发时跑，且标了 `continue-on-error`。
4. `release` —— 打 tag 且 arm64 成功时创建 Release。

打包步骤抽成了组合动作 `.github/actions/build-mac/action.yml`，两个架构共用一份。

> **Intel 为什么默认不发**：`macos-13` 在部分账号上会长时间排队（实测排队十几分钟不动），放进 tag 流程会把发版拖到 40 分钟超时。需要 x86_64 包时手动 Run workflow，跑完用 `gh release upload <tag> dist/PPT工具箱-mac.zip` 补到已发布的 Release 上。

> **`runs-on` 别写 `macos-14`**：实测该标签在本仓库会被解析成 `macos-14-arm64-us-xl`（XL 规格，需付费套餐），job 在分配机器前就被判掉，报 `The configuration 'macos-14-arm64-us-xl' is not supported`。`macos-15` / `macos-latest` 正常。
>
> 顺便记一条排查经验：run **0 秒失败、一个 job 都没创建**，说明工作流整份被判无效（常见于 job 级 `if` 里用了 `matrix`——它在矩阵展开前求值，拿不到 `matrix` 上下文）；run 有 job 但秒挂、`runner_name` 为空，则是 runner 标签不可用。

**额度**：公开仓库的 Actions 完全免费、不限分钟，所以这里不用为省钱做妥协。私有仓库则要小心——macOS runner 按 **10 倍**计费，2000 分钟免费额度实际只够约 200 分钟 macOS。

打包直接调用 `bash build_mac.sh`，等于顺带验证文档里那条命令没坏。打完不是上传就完事，还有一道**产物探针**：

```bash
dist/PPT工具箱.app/Contents/MacOS/PPT工具箱 --ci
```

这个 `--ci` 和给人看的 `--check` 不是一回事：它只挑「打包出来就该有」的硬指标（资源是否释放、用户数据目录是否可写、字体目录是否找到、内嵌自检样本是否还在、tkinter 是否可用），缺一项就退出非 0 让 CI 红掉。Office / LibreOffice 在构建机上本来就没有，所以不参与判定。

**发版**

```bash
git tag v1.08 && git push origin v1.08
```

Actions 会自动建 Release 并附上 arm64 的 zip。手工发版则直接把 `dist/PPT工具箱-mac.zip` 拖进去。

## 自检

不管哪个平台，遇到报错先做这一件事：**关于 → 绿色「一键功能自检」**。它会用内置的 3 份测试 PPT 依次验证资源释放、路径层、字体索引、字体检测、自动保存、图片化导出，并写出报告文件。报告里能直接看出是哪一环的路径不对，比盲试快得多。

命令行也可以：

```bash
python main.py --check      # 环境全量报告（人看）
python main.py --ci         # 打包产物探针（CI 用，有失败项则退出码 1）
python main.py --selftest   # 跑内置样本，端到端功能验证
```

## 项目结构

```
main.py            界面与入口（tkinter 自绘卡片，4 个导航页）
paths.py           跨平台路径层：Known Folder / 长路径 / 短路径 / 诊断
font_detect.py     OOXML 解析，提取所有字体引用并与本机字体比对
font_db.py         字体名与文件索引，带目录指纹缓存
font_install.py    字体安装（注册表 / atsutil）
slide_export.py    页面图片化（COM / LibreOffice / AppleScript 三条路）
autosave.py        定时自动保存与备份引擎
autostart.py       开机自启动（Run 注册表 / LaunchAgent）
selftest.py        六项功能自检
selftest_data.py   内嵌的 3 份测试 PPT（base64）
tools/             生成测试样本、重新内嵌资源的脚本
```

## 免责声明

- 本工具**只读**你选择的 PPT 文件；「自动保存」会写回原文件，「定时备份」写的是新副本，位置和份数都可控。
- 免费使用，**禁止倒卖与二次收费**；转载须保留作者与版本信息。
- 仅供学习和交流，请勿用于商业用途。

## 许可

见 [LICENSE](LICENSE)（自定义许可：可自由使用与修改，须保留署名，禁止商业销售）。
