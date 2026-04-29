#!/usr/bin/env python3
"""Read Project timeline.xlsx and write scripts/_project_timeline_grid.json for Smartsheet import."""
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path


def col_row(cell_ref: str):
    m = re.match(r"^([A-Z]+)(\d+)$", cell_ref)
    if not m:
        return None, None
    return m.group(1), int(m.group(2))


def col_to_idx(col: str) -> int:
    n = 0
    for c in col:
        n = n * 26 + (ord(c) - ord("A") + 1)
    return n - 1


def main():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else str(Path.home() / "Downloads" / "Project timeline.xlsx")
    out = sys.argv[2] if len(sys.argv) > 2 else "scripts/_project_timeline_grid.json"
    z = zipfile.ZipFile(xlsx)
    root = ET.fromstring(z.read("xl/worksheets/sheet1.xml"))
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ss = []
    if "xl/sharedStrings.xml" in z.namelist():
        sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
        for si in sroot.findall(".//m:si", ns):
            texts = []
            for t in si.findall(".//m:t", ns):
                if t.text:
                    texts.append(t.text)
            ss.append("".join(texts))

    rows: dict[int, dict[int, str]] = {}
    max_r = 0
    max_c = 0
    for row in root.findall(".//m:sheetData/m:row", ns):
        for c in row.findall("m:c", ns):
            ref = c.get("r")
            if not ref:
                continue
            col_letters, r = col_row(ref)
            if col_letters is None:
                continue
            ci = col_to_idx(col_letters)
            max_r = max(max_r, r)
            max_c = max(max_c, ci)
            t = c.get("t")
            v = c.find("m:v", ns)
            is_ = c.find("m:is", ns)
            val = None
            if t == "s" and v is not None and v.text is not None:
                val = ss[int(v.text)]
            elif v is not None and v.text is not None:
                val = v.text
                if t in (None, "n"):
                    try:
                        x = float(val)
                        if 30000 < x < 60000:
                            base = datetime(1899, 12, 30)
                            val = (base + timedelta(days=int(x))).strftime("%Y-%m-%d")
                    except (ValueError, OSError):
                        pass
            elif is_ is not None:
                tnodes = is_.findall(".//m:t", ns)
                val = "".join((t.text or "") for t in tnodes)
            rows.setdefault(r, {})[ci] = val if val is not None else ""

    grid = []
    for r in range(1, max_r + 1):
        row = []
        for ci in range(max_c + 1):
            row.append(rows.get(r, {}).get(ci, ""))
        grid.append(row)

    name = "Project timeline (from Excel)"
    with open(out, "w", encoding="utf-8") as f:
        json.dump({"name": name, "rows": grid}, f, indent=0)
    print(f"Wrote {out} ({len(grid)} rows x {len(grid[0])} cols)", file=sys.stderr)


if __name__ == "__main__":
    main()
