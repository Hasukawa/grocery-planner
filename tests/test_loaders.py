from pathlib import Path
import pytest

from ocado_tuesday import Item, load_tier1

XLSX = Path(__file__).parent.parent / "Ocado_Order_Manager.xlsx"


def test_load_tier1_returns_only_active_rows():
    items = load_tier1(XLSX)
    assert len(items) > 0
    for it in items:
        assert isinstance(it, Item)
        assert it.source == "tier1"
        assert it.name
        assert it.search_term
        assert it.qty >= 1


def test_load_tier1_first_item_is_apples():
    items = load_tier1(XLSX)
    assert items[0].name == "M&S Pink Lady Apples"
    assert items[0].search_term == "M&S Pink Lady Apples"
    assert items[0].qty == 1
