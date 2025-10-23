# 文档格式转换（AI + OCR）

支持将多种输入格式统一抽取为文本，经 LLM 重整后输出为目标格式：
- 输入：`docx` / `pdf` / `pptx` / `md` / `txt` / `html`
- 输出：`md` / `html` / `pptx` / `docx` / `pdf`

其中 `pdf` 在可复制文本时直接抽取；若为扫描版（无文本层），自动回退到 OCR（PaddleOCR + PyMuPDF）。

## 快速开始

1) 安装依赖（建议使用虚拟环境）

- 一次性安装（核心 + OCR）：

```
pip install -r requirements.txt
```

- 如 `paddlepaddle` 安装受限，可尝试使用镜像：

```
pip install paddlepaddle -i https://mirror.baidu.com/pypi/simple
```

2) 执行转换

```
python doc_transform.py --input <输入文件> --to <md|html|pptx|docx|pdf> --style concise
```

示例：

```
# 普通文档
python doc_transform.py --input document.md --to html --style concise

# 扫描版 PDF（需 OCR 依赖）
python doc_transform.py --input scanned_demo.pdf --to md --style concise

# 转换为 PDF 格式
python doc_transform.py --input document.docx --to pdf --style concise
python doc_transform.py --input presentation.pptx --to pdf --style concise
```

3) 生成测试用扫描 PDF（无文本层，仅图像）

```
python doc_transform.py --input scanned_demo.pdf --to md
```
