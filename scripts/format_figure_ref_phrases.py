"""Format complete '如图 x 所示' phrases without flattening native Word REF fields."""

from __future__ import annotations

import argparse
from copy import deepcopy
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree


W_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_URI}
Q = lambda name: f"{{{W_URI}}}{name}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
PATTERN = re.compile(r"如图\s*\d+\s*所示")


def set_size(run: etree._Element, half_points: int) -> None:
    rpr = run.find(Q("rPr"))
    if rpr is None:
        rpr = etree.Element(Q("rPr"))
        run.insert(0, rpr)
    for name in ("sz", "szCs"):
        element = rpr.find(Q(name))
        if element is None:
            element = etree.Element(Q(name))
            rpr.append(element)
        element.set(Q("val"), str(half_points))


def put_text(text_node: etree._Element, value: str) -> None:
    text_node.text = value
    if value.startswith(" ") or value.endswith(" "):
        text_node.set(XML_SPACE, "preserve")


def format_paragraph(paragraph: etree._Element, half_points: int) -> int:
    nodes = paragraph.xpath(".//w:t", namespaces=NS)
    content = "".join(node.text or "" for node in nodes)
    matches = list(PATTERN.finditer(content))
    if not matches:
        return 0
    selected = bytearray(len(content))
    for match in matches:
        selected[match.start():match.end()] = b"\x01" * (match.end() - match.start())

    offset = 0
    for node in nodes:
        original = node.text or ""
        length = len(original)
        marks = selected[offset:offset + length]
        offset += length
        if not any(marks):
            continue
        run = node.getparent()
        if run.tag != Q("r") or run.xpath("./w:t", namespaces=NS) != [node]:
            raise ValueError("Matched text is not in a single ordinary Word run")
        if any(child.tag not in (Q("rPr"), Q("t")) for child in run):
            raise ValueError("Matched run contains a field marker or object")
        if all(marks):
            set_size(run, half_points)
            continue

        parts = []
        start = 0
        while start < length:
            end = start + 1
            while end < length and marks[end] == marks[start]:
                end += 1
            parts.append((original[start:end], bool(marks[start])))
            start = end
        parent = run.getparent()
        position = parent.index(run)
        parent.remove(run)
        for value, chosen in parts:
            clone = deepcopy(run)
            put_text(clone.find(Q("t")), value)
            if chosen:
                set_size(clone, half_points)
            parent.insert(position, clone)
            position += 1
    return len(matches)


def format_docx(source: Path, output: Path, size_pt: float, expected_count: int) -> int:
    if source.resolve() == output.resolve() or output.exists():
        raise FileExistsError("Output must be a new path distinct from the source")
    half_points = round(size_pt * 2)
    if size_pt <= 0 or abs(size_pt * 2 - half_points) > 1e-8:
        raise ValueError("Word font size must be a positive multiple of 0.5 pt")
    original = source.read_bytes()
    with ZipFile(source) as zin:
        root = etree.fromstring(zin.read("word/document.xml"))
        paragraphs = root.xpath(".//w:body//w:p[not(ancestor::w:txbxContent)]", namespaces=NS)
        count = sum(format_paragraph(p, half_points) for p in paragraphs)
        if count != expected_count:
            raise ValueError(f"Expected {expected_count} phrases, found {count}; no output written")
        xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
        with ZipFile(output, "w") as zout:
            for item in zin.infolist():
                zout.writestr(deepcopy(item), xml if item.filename == "word/document.xml" else zin.read(item.filename))
    if source.read_bytes() != original:
        raise ValueError("Source changed unexpectedly")
    return count


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--size-pt", required=True, type=float)
    parser.add_argument("--expected-count", required=True, type=int)
    args = parser.parse_args()
    count = format_docx(args.source, args.output, args.size_pt, args.expected_count)
    print(f"Formatted {count} complete figure-reference phrases in {args.output}")


if __name__ == "__main__":
    main()
