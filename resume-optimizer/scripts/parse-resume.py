#!/usr/bin/env python3
"""解析简历文件（PDF / DOC / DOCX），输出纯文本到 stdout。"""

import sys
import os

def parse_pdf(path: str) -> str:
    try:
        import fitz  # pymupdf
    except ImportError:
        print("错误: 需要安装 pymupdf，请运行: pip install pymupdf", file=sys.stderr)
        sys.exit(1)
    doc = fitz.open(path)
    text = ""
    for page in doc:
        text += page.get_text()
    doc.close()
    return text

def parse_docx(path: str) -> str:
    try:
        from docx import Document
    except ImportError:
        print("错误: 需要安装 python-docx，请运行: pip install python-docx", file=sys.stderr)
        sys.exit(1)
    doc = Document(path)
    return "\n".join(p.text for p in doc.paragraphs)

def main():
    if len(sys.argv) < 2:
        print("用法: python parse-resume.py <简历文件路径>", file=sys.stderr)
        sys.exit(1)

    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"错误: 文件不存在: {path}", file=sys.stderr)
        sys.exit(1)

    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        text = parse_pdf(path)
    elif ext in (".doc", ".docx"):
        text = parse_docx(path)
    else:
        print(f"错误: 不支持的文件格式: {ext}（支持 .pdf .doc .docx）", file=sys.stderr)
        sys.exit(1)

    print(text)

if __name__ == "__main__":
    main()
