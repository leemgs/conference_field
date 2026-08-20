#!/usr/bin/env python3
"""Extract the yearly conference lists from the Samsung Research DOCX export."""

import json
import re
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "field-conference-list-v37-20260820_200532.docx"
OUTPUT = ROOT / "docs/data/conference_history.json"
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
VERSIONS = [
    ("2026-h2", "2026년 하반기", "2026-08-01부터 적용"),
    ("2026", "2026년", "2025-11-01부터 적용"),
    ("2025", "2025년", "2025년 컨퍼런스 리스트"),
    ("2024", "2024년", "2024년 컨퍼런스 리스트"),
]


def cell_text(cell):
    return "".join(node.text or "" for node in cell.findall(".//w:t", NS)).strip()


def clean_field(value):
    return re.sub(r"^\d{2}\.\s*", "", value).strip()


def extract():
    with ZipFile(SOURCE) as archive:
        root = ET.fromstring(archive.read("word/document.xml"))
    tables = root.findall(".//w:tbl", NS)
    if len(tables) != len(VERSIONS):
        raise ValueError(f"Expected {len(VERSIONS)} tables, found {len(tables)}")

    versions = []
    for table_index, (key, label, note) in enumerate(VERSIONS):
        table = tables[table_index]
        rows = []
        table_rows = table.findall("./w:tr", NS)[1:]
        for row_index, table_row in enumerate(table_rows, 1):
            cells = [cell_text(cell) for cell in table_row.findall("./w:tc", NS)]
            if table_index < 3:
                rating, abbreviation, title, h5 = (cells + [""] * 4)[:4]
                field = subfield = ""
            else:
                field, subfield, abbreviation, title, rating = (cells + [""] * 5)[:5]
                h5 = ""
            if not abbreviation or not title:
                continue
            rows.append({
                "id": f"{key}-{row_index:03d}",
                "field": clean_field(field),
                "subfield": subfield.strip(),
                "abbreviation": abbreviation.strip(),
                "title": title.strip(),
                "rating": rating.strip(),
                "h5Index": int(h5) if h5.isdigit() else None,
            })
        versions.append({"key": key, "label": label, "note": note, "conferences": rows})
    return {"source": SOURCE.name, "exported": "2026-08-20 20:05:31", "versions": versions}


if __name__ == "__main__":
    data = extract()
    OUTPUT.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("Wrote", OUTPUT, "(" + ", ".join(f"{v['label']}: {len(v['conferences'])}" for v in data["versions"]) + ")")
