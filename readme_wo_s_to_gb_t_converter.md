# WoS to GB/T Converter

This repository contains a small GUI Python utility that converts Web of Science (WoS) export spreadsheets into GB/T‑style reference entries and writes the results to a Word (`.docx`) file and an Excel file. The target workflow is double‑click → select file → get `参考文献输出.docx` and `参考文献输出.xlsx` in the same folder.

---

## Features

- GUI (no command line required) — double‑click the script and pick an input file.
- Supports common export formats: `.xlsx`, `.xls`, `.xlsb`, `.csv` (best effort; see notes).
- Intelligent author parsing:
  - Supports English names, comma formats ("Last, First"), and pinyin Chinese names (treated as surname‑first and abbreviated to initials).
  - Truncates authors to the first **3** by default and appends `et al.` (for non‑CJK) or `等.` (for CJK characters).
- Outputs references in a GB/T‑like layout using **English punctuation** for non‑CJK entries.
- Generates both a Word (`参考文献输出.docx`) and an Excel (`参考文献输出.xlsx`) file for review.

---

## Files

- `wos_to_gbt_gui_modern.py` — modern, recommended GUI script (Python 3.9+ compatible).
- `wos_to_gbt_gui_v3.py` — debugging‑friendly variant that writes an error log (`wos_to_gbt_error.log`).
- `demo_wos_export.xlsx` — small demo input used for testing (optional).

---

## Requirements

- **Python**: **3.9** (tested). The script should work with Python 3.8+ but has been developed and validated on 3.9.

### Python packages

Install the required packages before running:

```bash
pip install pandas openpyxl python-docx chardet
# Optional helpers for legacy/advanced formats
pip install pyxlsb pyexcel-xls
# If you must read legacy .xls files reliably, you can install:
# pip install xlrd==1.2.0
```

Notes:
- Modern `xlrd` (>=2.x) no longer supports `.xls`. If you need `.xls` support without converting files, install the legacy `xlrd==1.2.0`.
- `pyxlsb` enables `.xlsb` reading; `pyexcel-xls` is a fallback for some `.xls` cases.

---

## Quick Start (Double‑click GUI)

1. Place `wos_to_gbt_gui_modern.py` in a folder on your machine.
2. Install the dependencies shown above.
3. Double‑click `wos_to_gbt_gui_modern.py` (or run `python wos_to_gbt_gui_modern.py`).
4. In the file dialog, select your WoS export file (`.xlsx`, `.xls`, `.xlsb`, or `.csv`).
5. When the conversion completes, find the outputs in the same folder as the input:
   - `参考文献输出.docx`
   - `参考文献输出.xlsx`

---

## Command‑line (optional)

If you prefer to run from the command line (for automation), you can still run the script:

```bash
python wos_to_gbt_gui_modern.py
```

(The script will open the GUI file selector; it does not take CLI arguments in the current release.)

---

## Usage details & behavior

- Author parsing: The script detects CJK characters; if CJK characters are present it treats that author as Chinese and keeps the name as‑is. If author names are pinyin (Latin letters) but represent Chinese authors, the script uses a small pinyin surname list heuristic and formats `Surname Initials` (e.g. `Li Hongmei` → `Li H`).
- Fields used (attempted): `Publication Type`, `Authors`, `Article Title`, `Source Title`, `Volume`, `Issue`, `Article Number`, `Pages`, `DOI`, `Publication Year`. The script gracefully skips missing fields.
- Truncation: The default truncation is **3** authors. You can edit the script variable `truncate_n` to change this.

---

## Troubleshooting

- **Error: "File is not a zip file" when opening .xlsx**
  - This means the file extension is `.xlsx` but the content is not a valid `.xlsx` archive (often a CSV or corrupted file). Fixes:
    1. Open the file in Excel and **Save As** → `.xlsx`, then rerun.
    2. If the file is truly a `.csv`, choose the `.csv` file instead.

- **Error reading `.xls` files**
  - Modern `xlrd` no longer supports `.xls`. Install legacy support with `pip install xlrd==1.2.0`, or open the file in Excel and save as `.xlsx`.

- **Encoding issues with CSV**
  - If characters look garbled, try saving the CSV in UTF‑8 (or use a text editor to re-encode) or re-export from WoS with UTF‑8 encoding.

- **If conversion fails**
  - Use the debugging variant `wos_to_gbt_gui_v3.py` which writes `wos_to_gbt_error.log` containing a traceback. Attach that log when reporting issues.

---




---

# 中文说明

本仓库包含一个图形界面（GUI）Python 工具，用于将 Web of Science (WoS) 导出的参考文献表格转换为接近 GB/T 7714 风格的参考文献条目，并将结果写入 Word（`.docx`）和 Excel（`.xlsx`）文件。目标使用流程为：双击脚本 → 选择文件 → 在同一目录获得 `参考文献输出.docx` 和 `参考文献输出.xlsx`。

## 功能亮点

- 无需命令行，双击运行，弹出窗口选择文件。  
- 支持常见导出格式：`.xlsx`、`.xls`、`.xlsb`、`.csv`（视具体文件情况而定）。  
- 智能作者解析：支持英文名、逗号格式（"Last, First"）、拼音中文姓名（将拼音视为姓在前并取名首字母）。  
- 作者超过 3 人时默认截断并使用 `et al.`（非 CJK）或 `等.`（含 CJK 字符）。  
- 输出使用英文标点（英文/拼音文献），中文汉字作者保留原样并使用中文标点风格（在句内仍用英文标点以保持一致）。  
- 同时生成 Word 与 Excel，便于人工复核与二次处理。

## 文件说明

- `wos_to_gbt_gui_modern.py` — 推荐使用的现代版本脚本（适配 Python 3.9+）。
- `wos_to_gbt_gui_v3.py` — 调试友好版，遇到错误会写入 `wos_to_gbt_error.log`。
- `demo_wos_export.xlsx` — 可选的示例输入文件（非敏感数据）。

## 运行环境与依赖（安装示例）

建议使用 **Python 3.9**，也兼容 Python 3.8+。运行前请安装下列包：

```bash
pip install pandas openpyxl python-docx chardet
# 可选包（支持更多旧/二进制格式）
pip install pyxlsb pyexcel-xls
# 若需直接读取旧式 .xls 文件（不转换），可安装兼容旧版 xlrd：
# pip install xlrd==1.2.0
```

## 快速开始（双击 GUI）

1. 把 `wos_to_gbt_gui_modern.py` 放到你的电脑文件夹里。  
2. 安装依赖（见上）。  
3. 双击脚本（或在终端运行 `python wos_to_gbt_gui_modern.py`）。  
4. 在弹出的对话框中选择你的 WoS 导出文件（`.xlsx`、`.xls`、`.xlsb` 或 `.csv`）。  
5. 完成后在输入文件所在文件夹会生成： `参考文献输出.docx` 与 `参考文献输出.xlsx`。

## 常见问题及排查

- **错误“File is not a zip file”**：通常是文件扩展名为 `.xlsx`，但文件实际为 CSV 或已损坏。解决：在 Excel 中打开并另存为 `.xlsx`，或直接选择 CSV 文件。
- **无法读取 `.xls`**：现代 xlrd（>=2.0）不再支持 `.xls`。解决方法：安装旧版 `xlrd==1.2.0`，或在 Excel 中另存为 `.xlsx`。
- **CSV 编码问题**：出现中文乱码时请尝试用 UTF-8 编码保存 CSV 或在导出时选择 UTF-8 编码。
- **若转换失败**：使用调试版 `wos_to_gbt_gui_v3.py`，会在工作目录生成日志 `wos_to_gbt_error.log`，把该文件内容提供给维护者以便排查。


