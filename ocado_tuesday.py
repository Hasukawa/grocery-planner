"""Ocado Tuesday automation script."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

import openpyxl


@dataclass
class Item:
    name: str
    qty: int
    search_term: str
    notes: str
    source: Literal["tier1", "tier2"]


def load_tier1(xlsx_path: Path) -> list[Item]:
    """Load active rows from the 'Tier 1 – Essentials' sheet."""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Tier 1 – Essentials"]
    items: list[Item] = []
    # Row 1: title; row 2: headers; data from row 3
    for row in ws.iter_rows(min_row=3, values_only=True):
        name, _category, qty, search_term, notes, active = (row + (None,) * 6)[:6]
        if not name:
            continue
        if str(active or "").strip().lower() != "yes":
            continue
        items.append(
            Item(
                name=str(name).strip(),
                qty=int(qty or 1),
                search_term=str(search_term or name).strip(),
                notes=str(notes or "").strip(),
                source="tier1",
            )
        )
    return items
