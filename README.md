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


