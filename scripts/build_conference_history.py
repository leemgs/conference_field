#!/usr/bin/env python3
"""Extract the yearly conference lists from the field-industry DOCX export."""

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
FIELD_KEYS = {
    "System/OS": "sys", "Intelligence": "ai", "Data": "data",
    "Network & Communications": "net", "Security": "sec",
    "Programming & SE": "plse", "Human-Computer Interaction & Graphics": "hci",
    "Algorithms & Theory": "theory", "Hardware, Robotics & Electronics": "hw",
    "Health, Digital Health & Biometrics": "health", "Mechanics & Chemistry": "etc",
}
FIELD_NAMES = {key: name for name, key in FIELD_KEYS.items()}
FIELD_NAMES.update({"arvr": "AR/VR", "health": "Health, Digital Health & Biometrics", "etc": "Etc"})


def cell_text(cell):
    return "".join(node.text or "" for node in cell.findall(".//w:t", NS)).strip()


def clean_field(value):
    return re.sub(r"^\d{2}\.\s*", "", value).strip()


def normalized(value):
    return re.sub(r"[^a-z0-9]", "", value.lower())


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
                "fieldKey": FIELD_KEYS.get(clean_field(field), ""),
                "subfield": subfield.strip(),
                "abbreviation": abbreviation.strip(),
                "title": title.strip(),
                "rating": rating.strip(),
                "h5Index": int(h5) if h5.isdigit() else None,
            })
        versions.append({"key": key, "label": label, "note": note, "conferences": rows})

    # The 2025+ tables omit taxonomy columns. Reuse the 2024 taxonomy and the
    # current conference dataset so year filtering remains useful in the UI.
    by_abbreviation = {}
    by_title = {}
    for row in versions[-1]["conferences"]:
        for alias in re.split(r"[/&]", row["abbreviation"]):
            if normalized(alias):
                by_abbreviation[normalized(alias)] = row["fieldKey"]
        by_title[normalized(row["title"])] = row["fieldKey"]
    current_path = ROOT / "docs/data/conferences.json"
    if current_path.exists():
        current = json.loads(current_path.read_text(encoding="utf-8"))
        for conference in current.get("conferences", []):
            key = conference.get("field", "")
            by_abbreviation[normalized(re.sub(r"\s+20\d{2}$", "", conference.get("name", "")))] = key
            by_title[normalized(conference.get("fullName", ""))] = key
    for version in versions[:-1]:
        for row in version["conferences"]:
            aliases = [normalized(alias) for alias in re.split(r"[/&]", row["abbreviation"])]
            key = next((by_abbreviation.get(alias) for alias in aliases if by_abbreviation.get(alias)), "")
            key = key or by_title.get(normalized(row["title"]), "")
            row["fieldKey"] = key or "etc"
            row["field"] = FIELD_NAMES.get(row["fieldKey"], "Etc")
    return {"source": SOURCE.name, "exported": "2026-08-20 20:05:31", "versions": versions}


if __name__ == "__main__":
    data = extract()
    OUTPUT.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print("Wrote", OUTPUT, "(" + ", ".join(f"{v['label']}: {len(v['conferences'])}" for v in data["versions"]) + ")")
