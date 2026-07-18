#!/usr/bin/env python3
"""Audit chapter numbering and cross-reference fields in an academic DOCX."""

from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document


CAPTION_RE = re.compile(r"^(图|表)(\d+)-(\d+)")
EQUATION_RE = re.compile(r"\((\d+)-(\d+)\)\s*$")


def field_codes(document: Document) -> list[str]:
    codes: list[str] = []
    for node in document.element.body.iter():
        if node.tag.endswith("instrText") and node.text:
            codes.append(" ".join(node.text.split()))
    return codes


def bookmark_names(document: Document) -> set[str]:
    names: set[str] = set()
    for node in document.element.body.iter():
        if node.tag.endswith("bookmarkStart"):
            for key, value in node.attrib.items():
                if key.endswith("}name"):
                    names.add(value)
    return names


def audit(path: Path) -> list[str]:
    document = Document(path)
    codes = field_codes(document)
    bookmarks = bookmark_names(document)
    paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
    errors: list[str] = []

    required_fields = ("SEQ Figure", "SEQ Table", "SEQ Equation", "SEQ Reference", "REF ")
    for required in required_fields:
        if not any(required in code for code in codes):
            errors.append(f"缺少 Word 域：{required}")

    captions = [text for text in paragraphs if text.startswith(("图", "表"))]
    malformed_captions = [
        text for text in captions
        if ("Fig." in text or "Table" in text) and not CAPTION_RE.match(text)
    ]
    if malformed_captions:
        errors.append("存在非“章节-序号”格式的图题或表题：" + "；".join(malformed_captions[:3]))

    equations = [text for text in paragraphs if EQUATION_RE.search(text)]
    if equations and not any("SEQ Equation" in code for code in codes):
        errors.append("检测到公式编号文本，但没有 SEQ Equation 域")

    expected_prefixes = ("fig_", "tab_", "eq_", "ref_")
    for prefix in expected_prefixes:
        if not any(name.startswith(prefix) for name in bookmarks):
            errors.append(f"缺少 {prefix} 开头的交叉引用书签")

    ref_fields = [code for code in codes if code.startswith("REF ")]
    if not ref_fields:
        errors.append("没有检测到 REF 交叉引用域")

    return errors


def main() -> int:
    if len(sys.argv) != 2:
        print("用法：python audit_academic_docx.py <document.docx>")
        return 2
    path = Path(sys.argv[1])
    if not path.is_file():
        print(f"文件不存在：{path}")
        return 2
    errors = audit(path)
    if errors:
        print("[FAIL] 学术 Word 规范审计未通过：")
        for error in errors:
            print(f"- {error}")
        return 1
    print("[OK] 已检测到章节式图表/公式编号、参考文献序号、书签和 REF 交叉引用域。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
