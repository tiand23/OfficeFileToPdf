# OfficeFileToPdf

[中文](README.md) | [日本語](README.ja.md) | [English](README.en.md)

## 项目简介

这是一个仅面向 Windows 的批量转换工具。

它会递归扫描指定目录，把支持的 Office 文档和 PDF 统一输出到 `res/` 目录中，并尽量保持原有目录结构不变。

输出后的文件全部是 PDF。

同时，程序会在 `res/` 根目录生成一个 `source_to_pdf_map.json` 文件，用来记录源文件与新生成 PDF 的对应关系。

## 主要功能

- 递归扫描文件夹
- 保持原始目录结构输出到 `res/`
- Excel、Word、PowerPoint 文件转换为 PDF
- 已经是 PDF 的文件也会重新处理并输出到 `res/`
- 对 PDF 做自动去白边处理
- 生成源文件与目标 PDF 的映射 JSON
- 支持双击 `ToPdf.bat` 运行
- 支持把目标文件夹拖到 `ToPdf.bat` 上运行

## 运行前提

- 操作系统必须是 Windows
- 机器上必须安装 Microsoft Office
- Office 文档转换只使用 Microsoft Office COM
- 不使用 LibreOffice
- 首次运行时如果没有 Python，`ToPdf.bat` 会尝试自动安装 Python
- 首次运行时会自动创建 `.venv` 并安装 Python 依赖

## 支持的输入类型

支持转换为 PDF 的类型：

- Excel：`.xls` `.xlsx` `.xlsm` `.xlsb` `.csv` `.ods`
- Word：`.doc` `.docx` `.docm` `.rtf` `.odt`
- PowerPoint：`.ppt` `.pptx` `.pptm` `.odp`
- PDF：`.pdf`

不会转换的类型：

- 图片文件：`.png` `.jpg` `.jpeg` `.bmp` `.gif` `.svg` `.tif` `.tiff` `.webp` `.heic` `.ico`
- 文本文件：`.txt` `.md` `.json` `.yaml` `.yml` `.xml` `.toml` `.ini` `.cfg` `.log` `.rst`
- 工具脚本或程序文件：`.py` `.bat` `.cmd` `.ps1` `.exe` `.dll` `.sh`
- 无扩展名文件
- 其他暂不支持的文件类型

## 输出规则

- 输出目录默认是目标目录下的 `res/`
- 输出目录结构尽量保持和源目录一致
- 输出文件统一为 `.pdf`
- 如果源目录里本来就存在同名目标冲突，程序会自动改名，避免覆盖

示例：

源目录：

```text
input/
  report.xlsx
  deck.pptx
  docs/
    plan.docx
    old.pdf
    note.txt
```

输出目录：

```text
input/
  res/
    report.pdf
    deck.pdf
    docs/
      plan.pdf
      old.pdf
    source_to_pdf_map.json
```

说明：

- `note.txt` 会被跳过
- `old.pdf` 会重新输出到 `res/docs/old.pdf`
- `source_to_pdf_map.json` 会记录源文件与结果文件的关系

## 如何运行

### 方式 1：直接双击

直接双击 `ToPdf.bat`。

默认会处理 `ToPdf.bat` 所在目录，并在同级生成 `res/`。

### 方式 2：拖拽文件夹

把需要处理的文件夹直接拖到 `ToPdf.bat` 上。

程序会处理你拖进去的那个目录。

### 方式 3：命令行运行

在 Windows `cmd` 中执行：

```bat
cd /d D:\OfficeFileToPdf
ToPdf.bat "D:\YourInputFolder"
```

如果要处理脚本所在目录：

```bat
cd /d D:\OfficeFileToPdf
ToPdf.bat
```

## 映射 JSON 说明

程序会在 `res/source_to_pdf_map.json` 生成映射文件。

示例结构：

```json
{
  "source_root": "D:/input",
  "output_root": "D:/input/res",
  "mappings": [
    {
      "source_path": "D:/input/docs/plan.docx",
      "source_relative_path": "docs/plan.docx",
      "output_pdf_path": "D:/input/res/docs/plan.pdf",
      "output_relative_path": "docs/plan.pdf",
      "kind": "word"
    }
  ]
}
```

字段说明：

- `source_root`：源目录
- `output_root`：输出目录
- `mappings`：成功生成的映射列表
- `source_path`：源文件绝对路径
- `source_relative_path`：源文件相对路径
- `output_pdf_path`：生成后的 PDF 绝对路径
- `output_relative_path`：生成后的 PDF 相对 `res/` 的路径
- `kind`：文件类型，可能是 `excel`、`word`、`powerpoint`、`pdf`

## 日志文件

运行时会生成两个日志文件：

- `ToPdf_run.log`
- `ToPdf_python.log`

用途：

- `ToPdf_run.log`：记录批处理入口执行过程
- `ToPdf_python.log`：记录 Python 主脚本输出和报错

如果运行失败，请优先查看这两个日志。

## 注意事项

- 这是一个 Windows 专用工具
- Office 文档转换依赖 Microsoft Office 是否可正常打开该文件
- 如果某些文件格式虽然在扩展名列表中，但 Microsoft Office 无法正常打开，转换仍然会失败
- PDF 去白边失败时，会回退为直接复制 PDF，避免整批任务中断
- 跳过的文件不会出现在映射 JSON 的 `mappings` 中
- 失败文件目前也不会写入映射 JSON

## 常见问题

### 双击后窗口一闪而过

请使用最新版本的 `ToPdf.bat`。

新版脚本会在结束时停住窗口，并输出日志文件位置。

### 提示无法启动 `Excel.Application`、`Word.Application` 或 `PowerPoint.Application`

通常表示：

- 没有安装 Microsoft Office
- Office 没有正确注册 COM
- Office 当前环境不可用

### 没有生成某些文件

请检查：

- 该文件是否属于跳过类型
- `ToPdf_python.log` 中是否出现转换失败信息
- `source_to_pdf_map.json` 中是否存在该文件的映射记录

## 文件说明

- `ToPdf.py`：主程序
- `ToPdf.bat`：Windows 一键启动脚本
- `requirements.txt`：Python 依赖
- `pdfToPng.py`：历史保留的 PDF 裁边脚本

## 许可证

本项目使用 MIT License。

详见 [`LICENSE`](LICENSE)。
