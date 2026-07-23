#!/usr/bin/env python3
"""Reusable MathML -> Word OMML helpers for academic DOCX builders."""

from __future__ import annotations

import argparse
from copy import deepcopy
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree


M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
MML_NS = "http://www.w3.org/1998/Math/MathML"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

GREEK = {
    "rho": "ρ",
    "mu": "μ",
    "eta": "η",
    "theta": "θ",
    "lambda": "λ",
    "Delta": "Δ",
    "sigma": "σ",
    "gamma": "γ",
    "omega": "ω",
    "alpha": "α",
    "beta": "β",
}


def locate_mml2omml_xsl(explicit: str | Path | None = None) -> Path:
    """Locate Microsoft's MathML-to-OMML stylesheet."""
    if explicit:
        path = Path(explicit)
        if path.is_file():
            return path
        raise FileNotFoundError(path)

    candidates = [
        Path(r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL"),
        Path(r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL"),
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate

    for base in (Path(r"C:\Program Files\Microsoft Office"), Path(r"C:\Program Files (x86)\Microsoft Office")):
        if base.exists():
            found = next(base.rglob("MML2OMML.XSL"), None)
            if found:
                return found
    raise FileNotFoundError("Microsoft Office MML2OMML.XSL was not found")


def simple_identifier_mathml(
    base: str,
    *,
    sub: str | None = None,
    sup: str | None = None,
    upright_subscript: bool = True,
) -> str:
    """Build MathML for a variable with optional subscript/superscript."""
    base = GREEK.get(base, base)
    base_xml = f"<mi>{escape(base)}</mi>"
    node = base_xml

    if sub is not None:
        parts = []
        for index, item in enumerate(sub.split(",")):
            if index:
                parts.append("<mo>,</mo>")
            item = GREEK.get(item, item)
            variant = ' mathvariant="normal"' if upright_subscript and len(item) > 1 else ""
            if item.isdigit():
                parts.append(f"<mn>{escape(item)}</mn>")
            else:
                parts.append(f"<mi{variant}>{escape(item)}</mi>")
        node = f"<msub>{node}<mrow>{''.join(parts)}</mrow></msub>"

    if sup is not None:
        if sup.lstrip("-").isdigit():
            sup_xml = f"<mn>{escape(sup)}</mn>"
        else:
            sup_xml = f"<mi>{escape(GREEK.get(sup, sup))}</mi>"
        node = f"<msup>{node}{sup_xml}</msup>"

    return f'<math xmlns="{MML_NS}">{node}</math>'


def mathml_to_omml(
    mathml: str | bytes | etree._Element,
    *,
    inline: bool = True,
    xsl_path: str | Path | None = None,
) -> etree._Element:
    """Transform MathML into an inline oMath or display oMathPara element."""
    if isinstance(mathml, etree._Element):
        source = deepcopy(mathml)
    else:
        source = etree.fromstring(mathml.encode("utf-8") if isinstance(mathml, str) else mathml)

    transform = etree.XSLT(etree.parse(str(locate_mml2omml_xsl(xsl_path))))
    result = transform(source).getroot()

    if inline:
        if etree.QName(result).localname == "oMath":
            return deepcopy(result)
        found = result.find(f"{{{M_NS}}}oMath")
        if found is None:
            raise ValueError("The transformation did not produce an inline m:oMath node")
        return deepcopy(found)

    if etree.QName(result).localname == "oMathPara":
        return deepcopy(result)
    wrapper = etree.Element(f"{{{M_NS}}}oMathPara", nsmap={"m": M_NS})
    wrapper.append(deepcopy(result))
    return wrapper


def set_math_size(omml: etree._Element, points: float) -> etree._Element:
    """Set OMML run size in half-points and return the same element."""
    half_points = str(int(round(points * 2)))
    for math_run in omml.iter(f"{{{M_NS}}}r"):
        run_properties = math_run.find(f"{{{M_NS}}}rPr")
        if run_properties is None:
            run_properties = etree.Element(f"{{{M_NS}}}rPr")
            math_run.insert(0, run_properties)
        for local in ("sz", "szCs"):
            size = run_properties.find(f"{{{W_NS}}}{local}")
            if size is None:
                size = etree.SubElement(run_properties, f"{{{W_NS}}}{local}")
            size.set(f"{{{W_NS}}}val", half_points)
    return omml


def append_inline_math(paragraph, mathml: str, *, points: float | None = None, xsl_path=None):
    """Append a real inline OMML object to a python-docx paragraph."""
    omml = mathml_to_omml(mathml, inline=True, xsl_path=xsl_path)
    if points is not None:
        set_math_size(omml, points)
    paragraph._p.append(omml)
    return omml


def self_test() -> None:
    mathml = simple_identifier_mathml("P", sub="EH")
    inline = mathml_to_omml(mathml, inline=True)
    display = mathml_to_omml(mathml, inline=False)
    assert etree.QName(inline).localname == "oMath"
    assert etree.QName(display).localname == "oMathPara"
    text = "".join(inline.itertext())
    assert "P" in text and "EH" in text


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true", help="run a conversion smoke test")
    args = parser.parse_args()
    if not args.self_test:
        parser.error("use --self-test, or import this module from a DOCX builder")
    self_test()
    print("[OK] MathML -> inline/display OMML conversion passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

