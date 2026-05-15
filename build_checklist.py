"""Generate index.html for the Tier 2 checklist web app.

Reads the Rotation Pool sheet from the local xlsx and bakes its items into
the HTML template as a JavaScript array. Run this whenever you add/edit/remove
Rotation Pool items, then commit & push to update the deployed page.
"""
from __future__ import annotations

import json
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
XLSX = ROOT / "Ocado_Order_Manager.xlsx"
TEMPLATE = ROOT / "checklist_template.html"
OUTPUT = ROOT / "index.html"


def load_rotation_pool() -> list[dict]:
    wb = openpyxl.load_workbook(XLSX, data_only=True)
    ws = wb["Rotation Pool"]
    items: list[dict] = []
    # Row 1: title; row 2: headers; data from row 3.
    # Columns: Item Name, Category, Tier, Qty, Ocado Search Term, Last Ordered, Notes
    for row in ws.iter_rows(min_row=3, values_only=True):
        padded = (row + (None,) * 7)[:7]
        name, category, tier, qty, search_term, _last_ordered, notes = padded
        if not name:
            continue
        items.append({
            "name": str(name).strip(),
            "category": str(category or "").strip() or "Other",
            "tier": int(tier or 2),
            "qty": int(qty or 1),
            "search_term": str(search_term or name).strip(),
            "notes": str(notes or "").strip(),
        })
    return items


def main() -> None:
    items = load_rotation_pool()
    template = TEMPLATE.read_text(encoding="utf-8")
    rendered = template.replace("__ITEMS_JSON__", json.dumps(items, indent=2))
    OUTPUT.write_text(rendered, encoding="utf-8")
    print(f"Generated {OUTPUT.name} with {len(items)} items from Rotation Pool.")


if __name__ == "__main__":
    main()
