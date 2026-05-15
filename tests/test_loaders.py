from datetime import date
from pathlib import Path
import pytest

from ocado_tuesday import Item, load_tier1, load_tier2, upcoming_wednesday

XLSX = Path(__file__).parent.parent / "Ocado_Order_Manager.xlsx"
FIXTURE = Path(__file__).parent / "fixtures" / "tier2_sample.json"


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


def test_upcoming_wednesday_from_tuesday():
    # 2026-05-12 is a Tuesday; upcoming Wed is 2026-05-13
    assert upcoming_wednesday(date(2026, 5, 12)) == date(2026, 5, 13)


def test_upcoming_wednesday_from_wednesday_is_same_day():
    assert upcoming_wednesday(date(2026, 5, 13)) == date(2026, 5, 13)


def test_upcoming_wednesday_from_thursday_is_next_week():
    assert upcoming_wednesday(date(2026, 5, 14)) == date(2026, 5, 20)


def test_load_tier2_parses_items():
    items, week_str = load_tier2(FIXTURE)
    assert week_str == "2099-01-07"
    assert len(items) == 2
    assert items[0].name == "M&S Chicken Kyivs"
    assert items[0].qty == 2
    assert items[0].source == "tier2"


def test_load_tier2_missing_file_raises():
    with pytest.raises(FileNotFoundError):
        load_tier2(Path("/no/such/file.json"))
