# Batter Office 2016

<div align="center">
<img alt="Windows x64" src="https://img.shields.io/badge/Windows-x64-0078D4">
<img alt="Office 2016" src="https://img.shields.io/badge/Office-2016-D83B01">
<img alt="版本 Turbo" src="https://img.shields.io/badge/版本-Turbo-22863A">
<img alt="C%23 与 Python" src="https://img.shields.io/badge/技术-C%23%20%2B%20Python-512BD4">
</div>

**让 Word 和 PowerPoint 的 Ctrl+V 自动把数学公式转成可编辑的 Office 原生公式。**
**请下载 GitHub 仓库中的最新 ZIP 包。我会在使用过程中持续修复可能存在的 Bug。再次提醒：本补丁仅适用于 Office 2016，其他版本可借助 AI 工具对本补丁进行适配。**
从网页、ChatGPT 回答或 LaTeX 文本复制内容，再粘贴到 Office，正文与公式可以一起转换。适合整理学习笔记、数学推导和演示文稿。

> [!NOTE]
> **我本人使用的版本是 Microsoft Office 专业增强版 2016。**
> Word 和 PowerPoint 的日常复制粘贴已在我的电脑上测试通过。其他 Office 版本尚未逐一验证，欢迎反馈实际使用情况。
>
> 最新的 `aligned` 修复已通过程序层面的混合正文与公式转换检查，更多网页选区仍需要实际验证。

## 目录

- [功能](#功能)
- [运行环境](#运行环境)
- [安装与更新](#安装与更新)
- [使用示例](#使用示例)
- [还原与清理](#还原与清理)
- [速度与工作原理](#速度与工作原理)
- [使用范围与限制](#使用范围与限制)
- [常见问题](#常见问题)
- [反馈与贡献](#反馈与贡献)
- [致谢与许可](#致谢与许可)

## 功能

- **原来的快捷键就够了**，使用 Ctrl+C、Ctrl+V 完成复制和公式转换。
- **输出可编辑公式**，支持在 Office 公式编辑器中继续修改。
- **正文与公式一起粘贴**，保留中文、段落以及部分加粗、斜体样式。
- 识别 `\(...\)`、`\[...\]`、`$...$` 和 `$$...$$`。
- 从带 LaTeX 注释的 KaTeX 富文本、部分 MathML 内容中恢复公式。
- 支持常见分式、根号、上下标、求和以及多行 `aligned` 公式。
- 关闭求和号随嵌套内容自动拉伸的行为，改善多重求和的显示大小。
- 提供中文安装菜单，以及安装前加载项注册状态的备份与还原。

## 运行环境

| 项目 | 当前要求或验证情况 |
| --- | --- |
| 系统 | 64 位 Windows，安装包内置 x64 Python 运行环境 |
| Office | 本人使用 Microsoft Office 专业增强版 2016 |
| 目标应用 | Windows 桌面版 Word、PowerPoint |
| 转换组件 | 使用本机 Office 的 `MML2OMML.XSL` |
| .NET | 加载项使用 .NET Framework 4.x |
| 权限 | 安装到当前用户，通常无需管理员权限 |
| Python | 安装包已携带运行环境和转换依赖 |

当前程序会在常见的 Office16 安装目录查找转换组件。特殊安装路径可能需要适配。

## 安装与更新

1. 获取项目提供的完整 ZIP 安装包并解压。
2. 保存文档，完全退出 Word 和 PowerPoint。
3. 双击解压目录中的 `install.cmd`。
4. 输入 `1`，按 Enter 安装或更新。
5. 重新打开 Word 或 PowerPoint，正常复制、粘贴。

**使用普通用户权限即可。** 安装程序会把文件复制到当前用户的应用目录，并注册 COM 加载项。

```text
[1]  安装 / 更新
     启用 Word 和 PowerPoint 的 Ctrl+V 公式转换

[2]  还原
     恢复首次使用本安装包之前的加载项状态

[3]  退出
```

菜单会显示安装状态和 Office 运行状态。操作完成后按 Enter 返回菜单。

### CMD 命令

在解压目录中打开普通 CMD，可以直接运行：

```bat
install.cmd
install.cmd install
install.cmd restore
```

| 命令 | 用途 |
| --- | --- |
| `install.cmd` | 打开中文菜单 |
| `install.cmd install` | 安装或更新 |
| `install.cmd restore` | 恢复首次安装前的状态 |

更新时直接选择 **1**。首次安装前的备份会保留，后续更新沿用这份备份。

## 使用示例

安装后，在 Word 正文或 PowerPoint 幻灯片文本框中粘贴以下内容：

```latex
这是一个公式：\(\frac{a}{b}+\sqrt{x}\)。
```

预期会显示中文正文、分式和根号，公式可以继续编辑。

### 多行公式与正文

```latex
解方程组，得：

\[
\begin{aligned}
a_0 &= 13.45966387 \\
a_1 &= -3.60530940 \\
a_2 &= 0.26757066
\end{aligned}
\]

因此，所求最小二乘二次拟合多项式为：
\[
P_2(x)=13.45966-3.60531x+0.267571x^2
\]
```

也可以直接选中网页中已经排版好的正文和公式，一起复制。能否恢复精确公式，取决于来源网页放入剪贴板的内容。

## 还原与清理

运行 `install.cmd`，选择 **2**，恢复首次使用安装包之前的加载项注册状态，并移除本安装包部署的程序与备份。

> [!IMPORTANT]
> 如果首次安装前已经使用过测试版，**还原会恢复测试版的注册信息**。请保留这些注册信息指向的旧目录，直到确定不再需要回退。

安装目录：

```text
%LOCALAPPDATA%\BatterOffice\OfficeMathPastePackage\
├── app\         当前运行的加载项与转换引擎
└── original\    首次安装前的注册信息备份
```

安装完成后可以删除旧 ZIP 和解压副本。建议保留最新安装包，方便更新或还原。

## 速度与工作原理

加载项启动时准备转换引擎，同一 Office 应用中的后续粘贴复用该进程，减少重复启动和加载模块的时间。

```text
网页富文本 / LaTeX
        ↓
识别公式并提取源码
        ↓
LaTeX → MathML → Office Math（OMML）
        ↓
Word 文档 / PowerPoint 文本框
```

| 组件 | 作用 |
| --- | --- |
| C# COM 加载项 | 接入 Office，读取剪贴板并插入转换结果 |
| Python | 识别、转换和生成临时文档 |
| latex2mathml | 将 LaTeX 转为 MathML |
| lxml 与 Office XSLT | 生成 OMML 公式结构 |
| 常驻转换进程 | 在多次粘贴间复用转换组件 |

开发时使用同一条三重求和公式进行了预热后的 20 次连续测试：

| 转换目标 | 转换耗时中位数 |
| --- | --- |
| Word | 约 6.2 ms |
| PowerPoint | 约 11.5 ms |

以上数据来自开发机的 Turbo 版测试，计时覆盖转换请求到临时文件生成。Office 实际插入、渲染以及首次初始化还会增加耗时，复杂内容的结果也会不同。

转换在本机完成。程序通过临时文件和本地进程管道处理粘贴内容。

## 使用范围与限制

- 当前主要面向 Word、PowerPoint，其他 Office 应用尚未适配。
- 一般需要明确的公式分隔符，或富文本里可提取的公式源码。没有分隔符的裸 LaTeX 暂不自动识别。
- 网页必须提供可恢复的文本、LaTeX 或 MathML。只有图片的公式不能通过此工具识别。
- LaTeX 支持范围受转换库限制。复杂宏、自定义命令和部分环境可能保留原文或转换失败。
- `aligned` 当前转换为两列对齐结构，复杂的多对齐列布局可能需要进一步适配。
- 保留的正文格式包括基本段落、加粗和斜体。网页图片、复杂表格和完整 CSS 排版尚未完整支持。
- PowerPoint 插入时会使用 Office 剪贴板，后续粘贴可能使用转换后的 PowerPoint 内容。
- macOS、网页版 Office 及其他 Office 版本尚未验证。

## 常见问题

<details>
<summary>安装提示 Word 或 PowerPoint 仍在运行</summary>

先保存文档并完全退出两个应用，再返回菜单操作。如果窗口已关闭但仍提示运行，可在任务管理器中检查是否还有相应进程。

</details>

<details>
<summary>Office 提示加载项出现问题，或者粘贴后还是公式代码</summary>

记录使用的 Office 版本和触发问题的原始公式。如果 Office 已禁用加载项，可以在“文件 → 选项 → 加载项”查看状态。更新安装前先退出 Office，重新打开后再次测试。

请在反馈中说明是整段复制失败，还是单独复制公式也失败。两种情况有助于定位剪贴板格式和公式解析问题。

</details>

<details>
<summary>第一次粘贴比后续粘贴慢</summary>

第一次可能需要等待转换引擎初始化。Turbo 版会复用已启动的引擎，后续粘贴通常更快。

</details>

<details>
<summary>下载安装包后，脚本被 Windows 阻止</summary>

先确认安装包来源可信。如果文件属性中有“解除锁定”，可以在解压前解除 ZIP 的下载锁定后重新解压。单位电脑可能有脚本或加载项管理策略，需要由管理员按组织要求处理。

</details>

## 反馈与贡献

欢迎提交可复现的问题。建议附上：

- Office 的完整版本名称，以及发生问题的应用。
- 来源网页类型和复制方式，例如整段选中、单独选中公式、复制代码。
- 一小段可公开的原始 LaTeX。
- 预期结果与实际粘贴效果。
- 使用的安装包名称。

对公式解析和正文格式的改进，请同时检查 Word、PowerPoint 的结果，并保留普通粘贴的行为。

## 致谢与许可

最初的使用体验来自 Obsidian 的 Math Paste Bridge，随后逐步实现 Office 中的 Ctrl+V 公式转换。项目开发与问题排查由作者和 Codex 协作完成。

公式转换使用 Python、latex2mathml、lxml，以及本机 Office 提供的数学公式转换组件。README 编写参考了 [github-readme-skill](https://github.com/Ricardo-Nima/github-readme-skill)。

本项目尚未确定统一的开源许可证。第三方组件遵循各自的许可证；项目许可证以仓库最终提供的 LICENSE 文件为准。

