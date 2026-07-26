#!/usr/bin/env python3
"""Audit math, tables, fields, bookmarks, images, and references in an academic DOCX."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"w": W, "m": M, "a": A}

CAPTION_RE = re.compile(r"^(图|表)\s*\d+-\d+")
LATEX_RE = re.compile(r"\\(?:frac|rho|mu|eta|theta|lambda|sqrt|sum|log|mathrm|text|begin|end)\b")
UNDERSCORE_MATH_RE = re.compile(
    r"(?<![A-Za-z0-9])(?:[A-Za-z][A-Za-z0-9]*|BER|SNR|PTE)_[A-Za-z0-9,]+"
    r"|(?<![A-Za-z0-9])_[A-Za-z0-9,]+"
    r"|[A-Za-z][A-Za-z0-9]*_,[A-Za-z0-9]+"
)
PLAIN_GREEK_RE = re.compile(r"[ρθηλΔ]|μ(?![WHAFCVΩms]\b)")
REF_TARGET_RE = re.compile(r"\bREF\s+([A-Za-z_][A-Za-z0-9_]*)")
CITATION_COMMA_RE = re.compile(r"\[\d+\]\s*[,，]\s*\[\d+\]")
REFERENCE_ENTRY_RE = re.compile(r"^\[\d+\]\s+\S")


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def attr(node: etree._Element | None, local: str, default=None):
    return default if node is None else node.get(qn(W, local), default)


def paragraph_plain_text(paragraph: etree._Element) -> str:
    chunks = []
    for text_node in paragraph.xpath(".//w:t", namespaces=NS):
        if any(etree.QName(ancestor).namespace == M for ancestor in text_node.iterancestors()):
            continue
        if any(etree.QName(ancestor).localname == "instrText" for ancestor in text_node.iterancestors()):
            continue
        chunks.append(text_node.text or "")
    return "".join(chunks)


def paragraph_is_reference(paragraph: etree._Element) -> bool:
    for bookmark in paragraph.xpath(".//w:bookmarkStart", namespaces=NS):
        if (bookmark.get(qn(W, "name")) or "").startswith("ref_"):
            return True
    codes = " ".join(paragraph.xpath(".//w:instrText/text()", namespaces=NS))
    return "SEQ Reference" in codes


def border_value(parent: etree._Element, edge: str):
    node = parent.find(f"w:tblBorders/w:{edge}", namespaces=NS)
    return None if node is None else attr(node, "val")


def audit_table(table: etree._Element, index: int, strict: bool):
    errors, warnings = [], []
    table_properties = table.find("w:tblPr", namespaces=NS)
    if table_properties is None:
        errors.append(f"表{index}缺少tblPr")
        return errors, warnings

    fills = []
    for shading in table.xpath(".//w:shd", namespaces=NS):
        fill = attr(shading, "fill", "auto")
        if fill not in {"auto", "nil", "FFFFFF", "ffffff"}:
            fills.append(fill)
    if fills:
        errors.append(f"表{index}存在底色：{sorted(set(fills))}")

    top, bottom = border_value(table_properties, "top"), border_value(table_properties, "bottom")
    if strict and top != "single":
        errors.append(f"表{index}缺少三线表顶线")
    if strict and bottom != "single":
        errors.append(f"表{index}缺少三线表底线")
    for edge in ("left", "right", "insideH", "insideV"):
        value = border_value(table_properties, edge)
        if value not in {None, "nil", "none"}:
            errors.append(f"表{index}存在不允许的{edge}边框：{value}")

    rows = table.findall("w:tr", namespaces=NS)
    if rows:
        header_bottom = rows[0].xpath(".//w:tcBorders/w:bottom[@w:val='single']", namespaces=NS)
        if strict and not header_bottom:
            errors.append(f"表{index}缺少表头分隔线")
        header_repeat = rows[0].find("w:trPr/w:tblHeader", namespaces=NS)
        if strict and header_repeat is None:
            warnings.append(f"表{index}未设置重复表头")

    for row_number, row in enumerate(rows[1:], 2):
        forbidden = row.xpath(
            ".//w:tcBorders/*[self::w:left or self::w:right or self::w:top or self::w:bottom or self::w:insideH or self::w:insideV][not(@w:val='nil') and not(@w:val='none')]",
            namespaces=NS,
        )
        if forbidden:
            errors.append(f"表{index}第{row_number}行存在表体网格线")
            break

    if strict and table.find("w:tblGrid", namespaces=NS) is None:
        warnings.append(f"表{index}缺少显式tblGrid列宽")
    return errors, warnings


def audit(path: Path, strict: bool = False) -> dict:
    errors: list[str] = []
    warnings: list[str] = []
    try:
        with ZipFile(path) as package:
            bad_member = package.testzip()
            if bad_member:
                errors.append(f"DOCX压缩包损坏：{bad_member}")
            document_xml = package.read("word/document.xml")
    except (BadZipFile, KeyError) as exc:
        return {"path": str(path), "errors": [f"无法读取DOCX：{exc}"], "warnings": []}

    root = etree.fromstring(document_xml)
    body = root.find("w:body", namespaces=NS)
    if body is None:
        errors.append("document.xml缺少w:body")
        return {"path": str(path), "errors": errors, "warnings": warnings}

    math_nodes = root.xpath(".//m:oMath", namespaces=NS)
    display_nodes = []
    for node in math_nodes:
        paragraph = next((a for a in node.iterancestors() if a.tag == qn(W, "p")), None)
        paragraph_codes = "" if paragraph is None else " ".join(
            paragraph.xpath(".//w:instrText/text()", namespaces=NS)
        )
        if any(etree.QName(a).localname == "oMathPara" for a in node.iterancestors()) or "SEQ Equation" in paragraph_codes:
            display_nodes.append(node)
    inline_count = len(math_nodes) - len(display_nodes)

    field_codes = [" ".join(text.split()) for text in root.xpath(".//w:instrText/text()", namespaces=NS) if text.strip()]
    seq_counts = {
        "Figure": sum("SEQ Figure" in code for code in field_codes),
        "Table": sum("SEQ Table" in code for code in field_codes),
        "Equation": sum("SEQ Equation" in code for code in field_codes),
        "Reference": sum("SEQ Reference" in code for code in field_codes),
    }
    ref_codes = [code for code in field_codes if re.search(r"\bREF\s+", code)]

    bookmark_list = [node.get(qn(W, "name")) for node in root.xpath(".//w:bookmarkStart", namespaces=NS)]
    bookmark_list = [name for name in bookmark_list if name]
    bookmarks = set(bookmark_list)
    duplicates = sorted({name for name in bookmark_list if bookmark_list.count(name) > 1})
    if duplicates:
        errors.append("书签名称重复：" + "、".join(duplicates[:10]))

    missing_targets = []
    for code in ref_codes:
        match = REF_TARGET_RE.search(code)
        if match and match.group(1) not in bookmarks:
            missing_targets.append(match.group(1))
    if missing_targets:
        errors.append("REF目标书签不存在：" + "、".join(sorted(set(missing_targets))))

    tables = root.xpath(".//w:tbl", namespaces=NS)
    images = root.xpath(".//w:drawing", namespaces=NS)
    for index, table in enumerate(tables, 1):
        table_errors, table_warnings = audit_table(table, index, strict)
        errors.extend(table_errors)
        warnings.extend(table_warnings)

    plain_candidates = []
    latex_candidates = []
    citation_comma_groups = []
    for paragraph in body.xpath(".//w:p", namespaces=NS):
        if paragraph_is_reference(paragraph):
            continue
        text = paragraph_plain_text(paragraph)
        if not text:
            continue
        citation_comma_groups.extend(CITATION_COMMA_RE.findall(text))
        latex_candidates.extend(LATEX_RE.findall(text))
        plain_candidates.extend(UNDERSCORE_MATH_RE.findall(text))
        plain_candidates.extend(PLAIN_GREEK_RE.findall(text))
        if any(ancestor.tag == qn(W, "tbl") for ancestor in paragraph.iterancestors()):
            stripped = text.strip()
            if re.fullmatch(r"[dMkBUPRFINLCV]", stripped):
                plain_candidates.append(stripped)

    if latex_candidates:
        errors.append("存在残留LaTeX命令：" + "、".join(sorted(set(latex_candidates))))
    if strict and plain_candidates:
        errors.append("存在疑似普通文本数学变量：" + "、".join(sorted(set(plain_candidates))[:20]))
    if strict and not math_nodes:
        errors.append("严格模式下未检测到任何OMML数学对象")
    if citation_comma_groups:
        message = "多文献引用之间存在逗号：" + "、".join(
            sorted(set(citation_comma_groups))[:10]
        )
        (errors if strict else warnings).append(message)

    body_paragraph_texts = [
        paragraph_plain_text(paragraph).strip()
        for paragraph in body.xpath("./w:p", namespaces=NS)
    ]
    if "参考文献" in body_paragraph_texts:
        heading_index = body_paragraph_texts.index("参考文献")
        bibliography_entries = [
            text
            for text in body_paragraph_texts[heading_index + 1 :]
            if REFERENCE_ENTRY_RE.match(text)
        ]
        if not bibliography_entries:
            warnings.append(
                "参考文献区域为空：按默认策略保留空白并由作者填写，不自动生成或恢复题录"
            )

    if images and seq_counts["Figure"] == 0:
        errors.append("存在图片但缺少SEQ Figure域")
    if tables and seq_counts["Table"] == 0:
        errors.append("存在表格但缺少SEQ Table域")
    if display_nodes and seq_counts["Equation"] == 0:
        errors.append("存在独立公式但缺少SEQ Equation域")
    if seq_counts["Reference"] and not any(name.startswith("ref_") for name in bookmarks):
        errors.append("存在SEQ Reference但缺少ref_书签")
    if (seq_counts["Figure"] or seq_counts["Table"] or seq_counts["Equation"] or seq_counts["Reference"]) and not ref_codes:
        errors.append("存在自动编号但没有REF交叉引用域")

    captions = [paragraph_plain_text(p).strip() for p in body.xpath("./w:p", namespaces=NS)]
    malformed = [text for text in captions if text.startswith(("图", "表")) and not CAPTION_RE.match(text)]
    if malformed:
        warnings.append("疑似非章节式题注：" + "；".join(malformed[:3]))

    if strict and display_nodes and seq_counts["Equation"] != len(display_nodes):
        warnings.append(
            f"独立公式对象数({len(display_nodes)})与SEQ Equation数({seq_counts['Equation']})不同，请人工确认"
        )

    return {
        "path": str(path),
        "strict": strict,
        "omml_total": len(math_nodes),
        "omml_inline": inline_count,
        "omml_display": len(display_nodes),
        "tables": len(tables),
        "images": len(images),
        "seq_fields": seq_counts,
        "ref_fields": len(ref_codes),
        "bookmarks": len(bookmarks),
        "errors": errors,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("document", type=Path)
    parser.add_argument("--strict", action="store_true", help="enforce OMML and three-line-table rules")
    parser.add_argument("--json", action="store_true", help="print a machine-readable report")
    args = parser.parse_args()

    if not args.document.is_file():
        print(f"文件不存在：{args.document}", file=sys.stderr)
        return 2

    report = audit(args.document, strict=args.strict)
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(
            f"OMML={report.get('omml_total', 0)} "
            f"(inline={report.get('omml_inline', 0)}, display={report.get('omml_display', 0)}), "
            f"tables={report.get('tables', 0)}, images={report.get('images', 0)}, "
            f"REF={report.get('ref_fields', 0)}"
        )
        for warning in report.get("warnings", []):
            print(f"[WARN] {warning}")
        if report.get("errors"):
            print("[FAIL] 学术DOCX审计未通过：")
            for error in report["errors"]:
                print(f"- {error}")
        else:
            print("[OK] OMML、三线表、SEQ/REF、书签和图片结构审计通过。")
    return 1 if report.get("errors") else 0


if __name__ == "__main__":
    raise SystemExit(main())
