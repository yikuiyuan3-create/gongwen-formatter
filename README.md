# 公文一键排版 · WPS / Word 工具栏插件

[![检查](https://github.com/yikuiyuan3-create/gongwen-formatter/actions/workflows/verify.yml/badge.svg)](https://github.com/yikuiyuan3-create/gongwen-formatter/actions/workflows/verify.yml)
[![许可 MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

免费开源的 Windows DOCX 排版工具。在支持 COM 加载项的 **WPS 文字 / Microsoft Word** 中增加“公文排版”选项卡：点击“排版当前文档”，在排版窗口识别结构、按需编辑并导出新 DOCX。也可以从桌面打开，直接拖入文件。

**[下载安装包与校验值](https://github.com/yikuiyuan3-create/gongwen-formatter/releases/tag/v0.4.0)** · [安装与使用](docs/插件安装与使用.md) · [实机验收清单](docs/Windows实机验收清单.md)

当前为 **v0.4.0 测试版**。跨平台编译、核心检查、COM 激活与真正的 WPS/Word 工具栏操作是不同验证层。实际验证范围以 Release 说明和 Actions 结果为准，未经测试的宿主版本不视为已支持。

## 使用

1. 下载 Release 的 `公文一键排版_Setup_0.4.0_x64.exe`，关闭 WPS、Word 和旧版排版工具，再安装。
2. 重新打开 WPS 文字或 Word，将当前文件保存为 DOCX，并按 `Ctrl+S`。
3. 点击 **公文排版 → 排版当前文档**。工具会自动识别通知、请示、报告及常见段落角色。
4. 按需要编辑简单段落。编辑保护可仅在新副本中解除；普通自动编号的序列、缩进和制表位保留。
5. 点击 **4 导出 → 导出排好版 DOCX**，选择文件夹和新文件名。成功后显示完整路径，可打开文档或目录。

插件只读取已保存的当前文档，不替用户保存 WPS/Word 内容。工具已运行时接收导入请求，并保留换稿前未导出编辑提示。新建文档、旧 DOC、云端未落地文件需先另存为本机可读 DOCX。

## 功能与范围

- 本地离线处理，无需账号、API Key、订阅或按次付费。
- 识别标题、主送、正文、标题层级、附件说明、署名、日期；支持模板字体、字号、页边距和网格参数。
- 不覆盖原稿和已有输出；源文件变化会要求重新导入；仅明确编辑的简单段落会改变文本。
- 表格、图片、修订、字段和复杂结构保留；不安全 DOCX、宏、签名或不支持的复杂版式会被阻止。
- 本地审计只记录哈希、模板和处理元数据，不保存正文、文件名、路径。
- 不附商业字体，不破解打开密码，不处理 DOCM/扫描 PDF，不代表完整红头、版记、签章或正式发文审查。

## 环境

| 组件 | 要求与边界 |
| --- | --- |
| Windows | x64 Windows 10/11；建议仍受支持的 Windows 11。此包不适用于 Mac、Linux 或32位 Windows |
| 工具栏适配器 | .NET Framework 4.8 或更新4.x；AnyCPU，注册32/64位视图，可由不同位数办公软件加载，须分别实测 |
| WPS | 提供 COM 加载项和 Ribbon 接口的 Windows WPS 文字版本；不同渠道、版本和单位策略可能限制插件，不覆盖只有 JS 加载项的环境 |
| Word | Windows 桌面版、允许 COM 加载项的配置；不适用于网页版 Word |
| 排版窗口 | 自包含 .NET 10 Windows 运行组件，无需另外安装 .NET 10 |
| 字体 | 使用本机已安装且单位有权使用的字体；缺字体时明确提示 |

采用 **COM 工具栏入口 + 独立排版窗口**，不注入宏，不启用 HTTP 服务，不修改 Office 安全策略。宿主不允许 COM 时，可使用桌面拖入方式；这种情况不视为工具栏兼容验收通过。

## 开发与构建

需要 .NET 10 SDK、Python 3、NSIS 3（makensis）和7-Zip命令行（安装包校验）。首次恢复构建依赖需要联网，用户安装和排版可离线。

```sh
python3 scripts/create_samples.py
dotnet run --project tests/Gongwen.Tests -c Release -- .
dotnet run --project tests/Gongwen.OfficeAddin.Tests -c Release
python3 scripts/build.py --dotnet dotnet --runtime-version 10.0.13
python3 scripts/verify_installer.py
```

build.py 在新的版本目录发布；已有发布目录存在时停止，保留前次产物。安装器在 work/。工具栏适配器使用私有的 .NET Framework 参考程序集构建依赖，不需要开发者安装 Office/VSTO 或复制商业 PIA。

`--open <saved.docx>` 是工具栏调用入口，命令行和 IPC 严格校验路径，仅通过当前用户、本机会话命名管道转交，不自动排版或覆盖文件。

项目结构：Gongwen.Core 为 OOXML 引擎，Gongwen.Desktop 为 WPF 窗口，Gongwen.OfficeAddin 为 COM 工具栏；tests、packaging、scripts 分别为检查、安装器、构建脚本。

## 许可与反馈

采用 [MIT](LICENSE)，免费使用、修改、分发和商用。运行组件与厂商产品分别适用原有许可，见 [NOTICE.md](NOTICE.md)。

反馈请提供 Windows / WPS / Word 的版本与位数、脱敏步骤和虚构样例。不要上传真实业务文档或个人信息，见 [SECURITY.md](SECURITY.md)。
