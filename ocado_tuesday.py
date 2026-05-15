"""Ocado Tuesday automation script."""
from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import date, timedelta
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


def upcoming_wednesday(today: date) -> date:
    """The Wednesday on or after `today`."""
    # Monday=0 ... Wednesday=2 ... Sunday=6
    delta = (2 - today.weekday()) % 7
    return today + timedelta(days=delta)


def load_tier2(json_path: Path) -> tuple[list[Item], str]:
    """Return (items, week_string). Raises FileNotFoundError if missing."""
    if not json_path.exists():
        raise FileNotFoundError(json_path)
    with json_path.open() as f:
        data = json.load(f)
    week_str = str(data.get("week", "")).strip()
    items: list[Item] = []
    for entry in data.get("tier2_yes", []):
        name = str(entry.get("name", "")).strip()
        if not name:
            continue
        items.append(
            Item(
                name=name,
                qty=int(entry.get("qty", 1)),
                search_term=str(entry.get("search_term") or name).strip(),
                notes=str(entry.get("notes") or "").strip(),
                source="tier2",
            )
        )
    return items, week_str
