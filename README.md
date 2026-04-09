# WPSToLatex

> 将 WPS 表格（Spreadsheet）一键转换为 LaTeX 表格代码 | Convert WPS Spreadsheet tables to LaTeX table code with one click

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## 简介 | Description

**WPSToLatex** 是一个 WPS JS 宏工具，功能与知名的 [Excel2LaTeX](https://ctan.org/pkg/excel2latex) 插件类似。通过加载提供的 `.xlam` 宏文件，您可以在 WPS 表格中将选中的表格区域直接导出为 LaTeX `tabular` 代码，方便插入到 LaTeX 文档中。

**WPSToLatex** is a WPS JS macro tool, similar to the well-known [Excel2LaTeX](https://ctan.org/pkg/excel2latex) add-in. By loading the provided `.xlam` macro file, you can export selected table ranges in WPS Spreadsheets directly as LaTeX `tabular` code, ready to paste into your LaTeX documents.

---

## 功能特性 | Features

- ✅ 将 WPS 表格中的选区转换为 LaTeX `tabular` 环境代码
- ✅ 保留单元格合并（`\multirow`、`\multicolumn`）
- ✅ 支持边框线（`\hline`、`\cline`）
- ✅ 以对话框形式展示生成的 LaTeX 代码，可直接复制
- ✅ 基于 WPS JS 宏，无需安装额外运行时

---

## 安装与使用 | Installation & Usage

### 第一步：下载宏文件 | Step 1 – Download the macro file

从本仓库的 [Releases](../../releases) 页面下载最新的 `WPSToLatex.xlam` 文件，保存到本地任意目录。

Download the latest `WPSToLatex.xlam` from the [Releases](../../releases) page and save it to a local directory.

---

### 第二步：在 WPS 中加载宏文件 | Step 2 – Load the macro file in WPS

1. 用 **WPS 表格**（WPS Spreadsheets）打开给定的 `.xlam` 宏文件。
2. 点击顶部菜单 **工具（Tools）** → **选项（Options）**。
3. 在弹出的「选项」对话框中，选择左侧的 **自定义功能区（Customize Ribbon）** 选项卡。
4. 点击对话框上方 **从下列位置选择命令**，点击 **宏**，找到 **WPSToLatex**。
5. 在「自定义功能区」界面中，将 `WPSToLatex` 相关命令从左侧命令列表拖拽（或通过「添加>>」按钮）添加到右侧的目标选项卡 / 组中。
> 可以添加到已有选项卡中，也可以自己新建
6. 点击 **确定（OK）** 保存设置。

完成后，您将在 WPS 表格的功能区看到新增的 **WPSToLatex** 按钮。

---

### 第三步：使用宏转换表格 | Step 3 – Convert a table

1. 在 WPS 表格中，**选中**需要转换的表格区域（可包含合并单元格和边框）。
2. 点击功能区中的 **WPSToLatex** 按钮。
3. 弹出对话框，**LaTeX 已复制到剪贴板！文本长度：XXX**，成功！
4. 将代码粘贴到您的 `.tex` 文件中使用。

---

## 输出示例 | Output Example

选中以下 WPS 表格区域：

| A  | B  | C  |
|----|----|----|
| 1  | 2  | 3  |
| 4  | 5  | 6  |

生成的 LaTeX 代码示例：

```latex
\begin{tabular}{|c|c|c|}
\hline
A & B & C \\
\hline
1 & 2 & 3 \\
\hline
4 & 5 & 6 \\
\hline
\end{tabular}
```

---

## 测试平台 | Tested Platforms

WPS Office 2026 v12.1.2.25838（Zorin OS 18 based on Ubuntu 24.04, Linux）

---

## 许可证 | License

本项目基于 [MIT License](LICENSE) 开源。

This project is licensed under the [MIT License](LICENSE).
