# Ocado Tuesday Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a double-clickable macOS script that uses Playwright to fill an Ocado basket each Tuesday from a Tier 1 Excel sheet and a Tier 2 JSON file.

**Architecture:** Single Python script (`ocado_tuesday.py`) with focused functions for loading data, driving a persistent Chromium session, clearing the reserved order, and adding items. Selectors live as named constants at the top of the file for easy first-run tweaking. Loaders are TDD-tested; browser code is verified by running against the live site.

**Tech Stack:** Python 3.12, Playwright (sync API), openpyxl, pytest.

**Spec:** `docs/superpowers/specs/2026-05-14-ocado-tuesday-design.md`

---

## File structure

```
~/Code/grocery-planner/
├── ocado_tuesday.py            # main script
├── run_ocado.command           # macOS double-click launcher
├── requirements.txt
├── README.md
├── .gitignore
├── Ocado_Order_Manager.xlsx    (existing)
├── docs/superpowers/
│   ├── specs/2026-05-14-ocado-tuesday-design.md
│   └── plans/2026-05-14-ocado-tuesday.md  (this file)
├── tests/
│   ├── __init__.py
│   ├── test_loaders.py
│   └── fixtures/
│       └── tier2_sample.json
├── .ocado_session/             # auto-created at runtime
└── logs/                       # auto-created at runtime
```

`ocado_tuesday.py` keeps all runtime logic in one file. Tests cover only the deterministic loader functions; the Playwright code is verified by running the script against the live site (no realistic way to unit-test browser automation against a third-party DOM).

---

## Task 1: Project scaffolding

**Files:**
- Create: `.gitignore`
- Create: `requirements.txt`
- Create: `tests/__init__.py`
- Create: `logs/.gitkeep`

- [ ] **Step 1: Initialize git repo**

```bash
cd ~/Code/grocery-planner
git init
git branch -M main
```

- [ ] **Step 2: Write `.gitignore`**

```
.venv/
.ocado_session/
logs/*.log
__pycache__/
*.pyc
.DS_Store
.pytest_cache/
```

- [ ] **Step 3: Write `requirements.txt`**

```
playwright>=1.49,<2
openpyxl>=3.1,<4
pytest>=8.0,<9
```

- [ ] **Step 4: Create empty test package and logs placeholder**

```bash
mkdir -p tests/fixtures logs
touch tests/__init__.py logs/.gitkeep
```

- [ ] **Step 5: Set up venv and install**

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
playwright install chromium
```

Expected: chromium downloads, no errors.

- [ ] **Step 6: Commit**

```bash
git add .gitignore requirements.txt tests/__init__.py logs/.gitkeep
git commit -m "chore: scaffold grocery-planner project"
```

---

## Task 2: Item dataclass and Tier 1 loader (TDD)

**Files:**
- Create: `ocado_tuesday.py` (initial)
- Create: `tests/test_loaders.py`

The Excel file has a title row at row 1 and headers at row 2; data starts at row 3. Columns: `Item Name, Category, Qty, Ocado Search Term, Notes, Active?`. Only load rows where `Active?` equals `"Yes"` (case-insensitive).

- [ ] **Step 1: Write the failing test**

`tests/test_loaders.py`:

```python
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
```

- [ ] **Step 2: Run test, verify it fails**

```bash
source .venv/bin/activate
pytest tests/test_loaders.py -v
```

Expected: ImportError (module doesn't exist yet).

- [ ] **Step 3: Create minimal `ocado_tuesday.py` with Item + load_tier1**

```python
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
```

- [ ] **Step 4: Run tests, verify they pass**

```bash
pytest tests/test_loaders.py -v
```

Expected: both tests PASS.

- [ ] **Step 5: Commit**

```bash
git add ocado_tuesday.py tests/test_loaders.py
git commit -m "feat: add Item dataclass and Tier 1 loader"
```

---

## Task 3: Tier 2 loader (TDD)

**Files:**
- Modify: `ocado_tuesday.py` — append `load_tier2` and helpers
- Modify: `tests/test_loaders.py` — append Tier 2 tests
- Create: `tests/fixtures/tier2_sample.json`

The JSON format from the spec: `{"week": "YYYY-MM-DD", "tier2_yes": [{"name": "...", "qty": 1, "search_term": "..."}]}`.

Staleness rule: `week` must equal the upcoming Wednesday's date (the Wednesday on or after today). If mismatched, the loader signals stale; the caller prompts y/n. Missing file → raise `FileNotFoundError`.

- [ ] **Step 1: Create the fixture**

`tests/fixtures/tier2_sample.json`:

```json
{
  "week": "2099-01-07",
  "tier2_yes": [
    {"name": "M&S Chicken Kyivs", "qty": 2, "search_term": "M&S 2 Chicken Kyivs 320g"},
    {"name": "M&S Duck Breast", "qty": 1, "search_term": "M&S 2 Duck Breast Portions 265g"}
  ]
}
```

- [ ] **Step 2: Write failing tests**

Append to `tests/test_loaders.py`:

```python
from datetime import date
from ocado_tuesday import load_tier2, upcoming_wednesday

FIXTURE = Path(__file__).parent / "fixtures" / "tier2_sample.json"


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
```

- [ ] **Step 3: Run, verify failure**

```bash
pytest tests/test_loaders.py -v
```

Expected: ImportError on `load_tier2` / `upcoming_wednesday`.

- [ ] **Step 4: Implement in `ocado_tuesday.py`**

Add these imports near the top:

```python
import json
from datetime import date, timedelta
```

Add these functions after `load_tier1`:

```python
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
        name = entry.get("name", "").strip()
        if not name:
            continue
        items.append(
            Item(
                name=name,
                qty=int(entry.get("qty", 1)),
                search_term=entry.get("search_term", name).strip(),
                notes=entry.get("notes", "").strip(),
                source="tier2",
            )
        )
    return items, week_str
```

- [ ] **Step 5: Run, verify pass**

```bash
pytest tests/test_loaders.py -v
```

Expected: all 7 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add ocado_tuesday.py tests/test_loaders.py tests/fixtures/tier2_sample.json
git commit -m "feat: add Tier 2 JSON loader with week-staleness helper"
```

---

## Task 4: Constants, logging, and browser scaffolding

**Files:**
- Modify: `ocado_tuesday.py`

Set up all the constants in one block so first-run tweaks are easy. Add a logger that writes to both stdout and `logs/run-YYYY-MM-DD-HHMM.log`.

- [ ] **Step 1: Add constants and logging setup near the top of `ocado_tuesday.py`**

Place after the imports, before `Item`:

```python
import logging
import sys
from datetime import datetime

ROOT = Path(__file__).parent
XLSX_PATH = ROOT / "Ocado_Order_Manager.xlsx"
JSON_PATH = Path.home() / "Downloads" / "ocado_tuesday.json"
SESSION_DIR = ROOT / ".ocado_session"
LOGS_DIR = ROOT / "logs"

URL_RESERVED = "https://www.ocado.com/webshop/reservedOrder.do"
URL_HOME = "https://www.ocado.com/"
LOGIN_URL_FRAGMENTS = ("/login", "/signin")

# Selectors — comma-separated fallbacks. Tweak on first run if Ocado has changed.
SEL_SEARCH_INPUT = '[data-testid="search-input"], input[name="search"], input[type="search"]'
SEL_SEARCH_SUBMIT = 'button[type="submit"][aria-label*="earch"], button:has-text("Search")'
SEL_PRODUCT_CARD = '[data-testid^="product-tile"], article[data-product-id], li[data-product-id]'
SEL_PRODUCT_TITLE = 'h1, [data-testid="product-title"]'
SEL_ADD_BUTTON = 'button:has-text("Add"), button[data-testid*="add-to-trolley"]'
SEL_QTY_PLUS = 'button[aria-label*="ncrease"], button:has-text("+")'
SEL_REMOVE_ALL = 'button:has-text("Remove all"), button:has-text("Empty trolley"), button:has-text("Clear trolley")'
SEL_REMOVE_CONFIRM = 'button:has-text("Yes"), button:has-text("Confirm"), button:has-text("Remove")'
SEL_SIGN_IN = 'a:has-text("Sign in"), button:has-text("Sign in")'

ELEMENT_TIMEOUT_MS = 15_000
NAVIGATION_TIMEOUT_MS = 30_000


def setup_logging() -> Path:
    LOGS_DIR.mkdir(exist_ok=True)
    log_path = LOGS_DIR / f"run-{datetime.now():%Y-%m-%d-%H%M}.log"
    handler_file = logging.FileHandler(log_path)
    handler_stdout = logging.StreamHandler(sys.stdout)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s", datefmt="%H:%M:%S")
    for h in (handler_file, handler_stdout):
        h.setFormatter(fmt)
    logging.basicConfig(level=logging.INFO, handlers=[handler_file, handler_stdout])
    return log_path
```

- [ ] **Step 2: Add `ensure_logged_in` helper**

Append to `ocado_tuesday.py`:

```python
from playwright.sync_api import Page, TimeoutError as PWTimeout


def is_logged_out(page: Page) -> bool:
    url = page.url
    if any(frag in url for frag in LOGIN_URL_FRAGMENTS):
        return True
    sign_in = page.locator(SEL_SIGN_IN).first
    try:
        return sign_in.is_visible(timeout=2_000)
    except PWTimeout:
        return False


def ensure_logged_in(page: Page) -> None:
    log = logging.getLogger("ocado")
    page.goto(URL_RESERVED, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
    page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    for attempt in (1, 2):
        if not is_logged_out(page):
            log.info("Login check OK")
            return
        log.warning("Not logged in (attempt %d). Log in in the browser, then press Enter here.", attempt)
        input(">>> Press Enter once you've logged in: ")
        page.goto(URL_RESERVED, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
        page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    raise RuntimeError("Still not logged in after retry — aborting.")
```

- [ ] **Step 3: Verify the module still imports cleanly**

```bash
python -c "import ocado_tuesday; print('ok')"
```

Expected: prints `ok` with no errors.

- [ ] **Step 4: Commit**

```bash
git add ocado_tuesday.py
git commit -m "feat: add constants, logging, and login-check helper"
```

---

## Task 5: Clear reserved order

**Files:**
- Modify: `ocado_tuesday.py`

- [ ] **Step 1: Add `clear_reserved_order`**

Append to `ocado_tuesday.py`:

```python
def clear_reserved_order(page: Page) -> None:
    log = logging.getLogger("ocado")
    log.info("Clearing reserved order…")
    page.goto(URL_RESERVED, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
    page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    remove_btn = page.locator(SEL_REMOVE_ALL).first
    try:
        remove_btn.wait_for(state="visible", timeout=5_000)
    except PWTimeout:
        log.info("No 'Remove all' button found — assuming basket already empty.")
        return
    remove_btn.click()
    log.info("Clicked 'Remove all'.")
    confirm = page.locator(SEL_REMOVE_CONFIRM).first
    try:
        confirm.wait_for(state="visible", timeout=3_000)
        confirm.click()
        log.info("Confirmed removal.")
    except PWTimeout:
        log.info("No confirmation dialog appeared — that's fine.")
    page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
```

- [ ] **Step 2: Verify import still works**

```bash
python -c "import ocado_tuesday; print('ok')"
```

- [ ] **Step 3: Commit**

```bash
git add ocado_tuesday.py
git commit -m "feat: add reserved-order clearing"
```

---

## Task 6: add_item

**Files:**
- Modify: `ocado_tuesday.py`

Two code paths:
- **URL path** — `item.notes` starts with `https://www.ocado.com` → navigate directly
- **Search path** — fill search input, press Enter, wait for results, click first product card

On the product page: read title, fuzzy-match against `search_term`, log a warning on mismatch but continue. Click Add once, then "+" `(qty - 1)` more times.

- [ ] **Step 1: Add result type and add_item function**

Append to `ocado_tuesday.py`:

```python
@dataclass
class Result:
    item: Item
    status: Literal["ok", "not_found", "error"]
    detail: str = ""


def _titles_roughly_match(actual_title: str, search_term: str) -> bool:
    """True if the longest word in search_term appears (case-insensitive) in actual_title."""
    words = [w for w in search_term.split() if len(w) >= 4]
    if not words:
        return True
    longest = max(words, key=len)
    return longest.lower() in actual_title.lower()


def add_item(page: Page, item: Item) -> Result:
    log = logging.getLogger("ocado")
    log.info("→ %s [%s] qty=%d", item.name, item.source, item.qty)
    try:
        if item.notes.startswith("https://www.ocado.com"):
            page.goto(item.notes, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
        else:
            search_box = page.locator(SEL_SEARCH_INPUT).first
            search_box.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            search_box.fill(item.search_term)
            search_box.press("Enter")
            page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
            first_card = page.locator(SEL_PRODUCT_CARD).first
            try:
                first_card.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            except PWTimeout:
                return Result(item, "not_found", f"no results for '{item.search_term}'")
            first_card.click()

        title_loc = page.locator(SEL_PRODUCT_TITLE).first
        title_loc.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        actual = (title_loc.text_content() or "").strip()
        if not _titles_roughly_match(actual, item.search_term):
            log.warning("Title mismatch: wanted ~'%s', got '%s' — continuing", item.search_term, actual)

        add_btn = page.locator(SEL_ADD_BUTTON).first
        add_btn.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        add_btn.click()
        log.info("   added 1")

        for n in range(item.qty - 1):
            plus = page.locator(SEL_QTY_PLUS).first
            plus.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            plus.click()
            log.info("   qty +1 (now %d)", n + 2)

        return Result(item, "ok", actual)
    except PWTimeout as e:
        return Result(item, "error", f"timeout: {e}")
    except Exception as e:  # noqa: BLE001
        return Result(item, "error", f"{type(e).__name__}: {e}")
```

- [ ] **Step 2: Verify import**

```bash
python -c "import ocado_tuesday; print('ok')"
```

- [ ] **Step 3: Commit**

```bash
git add ocado_tuesday.py
git commit -m "feat: add per-item search/add flow"
```

---

## Task 7: Main flow and summary

**Files:**
- Modify: `ocado_tuesday.py`

Wires everything together: load data, check staleness, launch persistent context, login check, clear basket, add each item, print summary, wait for Enter to close.

- [ ] **Step 1: Add summary helper and main**

Append to `ocado_tuesday.py`:

```python
from playwright.sync_api import sync_playwright


def print_summary(results: list[Result], log_path: Path) -> None:
    tier1 = [r for r in results if r.item.source == "tier1"]
    tier2 = [r for r in results if r.item.source == "tier2"]
    ok1 = sum(1 for r in tier1 if r.status == "ok")
    ok2 = sum(1 for r in tier2 if r.status == "ok")
    not_found = [r for r in results if r.status == "not_found"]
    errors = [r for r in results if r.status == "error"]

    print()
    print("=" * 41)
    print("              SUMMARY")
    print("=" * 41)
    print(f"Tier 1: {ok1}/{len(tier1)} added")
    print(f"Tier 2: {ok2}/{len(tier2)} added")
    if not_found:
        print(f"\nNot found ({len(not_found)}):")
        for r in not_found:
            print(f"  - {r.item.name} [{r.item.source}]   search: \"{r.item.search_term}\"")
    if errors:
        print(f"\nErrors ({len(errors)}):")
        for r in errors:
            print(f"  - {r.item.name} [{r.item.source}]: {r.detail}")
    print(f"\nFull log: {log_path}")
    print("=" * 41)


def confirm_stale_week(week_str: str, expected: date) -> bool:
    print(f"\n⚠  Tier 2 JSON week is '{week_str}', expected '{expected.isoformat()}'.")
    answer = input("   Proceed anyway? [y/N]: ").strip().lower()
    return answer == "y"


def main() -> int:
    log_path = setup_logging()
    log = logging.getLogger("ocado")
    log.info("Run starting. Log: %s", log_path)

    tier1 = load_tier1(XLSX_PATH)
    log.info("Loaded %d Tier 1 items", len(tier1))

    try:
        tier2, week_str = load_tier2(JSON_PATH)
    except FileNotFoundError:
        log.error("Tier 2 JSON not found at %s — aborting.", JSON_PATH)
        return 1
    log.info("Loaded %d Tier 2 items (week=%s)", len(tier2), week_str)

    expected_week = upcoming_wednesday(date.today())
    if week_str != expected_week.isoformat():
        if not confirm_stale_week(week_str, expected_week):
            log.info("User declined stale JSON — aborting.")
            return 1

    all_items = tier1 + tier2
    results: list[Result] = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(SESSION_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        context.set_default_timeout(ELEMENT_TIMEOUT_MS)
        context.set_default_navigation_timeout(NAVIGATION_TIMEOUT_MS)
        page = context.pages[0] if context.pages else context.new_page()

        try:
            ensure_logged_in(page)
            clear_reserved_order(page)
            for item in all_items:
                results.append(add_item(page, item))
        except Exception as e:  # noqa: BLE001
            log.exception("Run aborted: %s", e)

        print_summary(results, log_path)
        input("\nBrowser left open for review. Press Enter to close.")
        context.close()

    return 0


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 2: Syntax check**

```bash
python -c "import ocado_tuesday; print('ok')"
```

- [ ] **Step 3: Commit**

```bash
git add ocado_tuesday.py
git commit -m "feat: wire up main flow and summary"
```

---

## Task 8: Launcher, README, final polish

**Files:**
- Create: `run_ocado.command`
- Create: `README.md`

- [ ] **Step 1: Write `run_ocado.command`**

```bash
#!/bin/bash
set -u
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "Virtualenv not found at .venv — run the setup steps in README.md first."
  echo
  read -n1 -s -r -p "Press any key to close."
  exit 1
fi
source .venv/bin/activate
python ocado_tuesday.py
status=$?
echo
read -n1 -s -r -p "Done (exit $status). Press any key to close."
```

- [ ] **Step 2: Make it executable**

```bash
chmod +x run_ocado.command
```

- [ ] **Step 3: Write `README.md`**

```markdown
# Ocado Tuesday

Weekly script that clears my reserved Ocado order and refills it from a Tier 1 Excel sheet plus a Tier 2 JSON checklist.

## First-run setup

1. `cd ~/Code/grocery-planner`
2. `python3 -m venv .venv`
3. `source .venv/bin/activate`
4. `pip install -r requirements.txt`
5. `playwright install chromium`
6. First run: `python ocado_tuesday.py`
   - A Chromium window opens. Log in to Ocado manually.
   - Switch back to the terminal and press Enter.
   - The script proceeds. The session is saved to `.ocado_session/` for future runs.

## Weekly run

1. Generate `~/Downloads/ocado_tuesday.json` from the checklist app. Format:
   ```json
   {
     "week": "2026-05-13",
     "tier2_yes": [
       {"name": "M&S Chicken Kyivs", "qty": 2, "search_term": "M&S 2 Chicken Kyivs 320g"}
     ]
   }
   ```
   `week` should be the upcoming Wednesday (delivery date).
2. Double-click `run_ocado.command`.

## Files

- `ocado_tuesday.py` — the script
- `Ocado_Order_Manager.xlsx` — Tier 1 essentials (edit the "Tier 1 – Essentials" sheet; only `Active? = Yes` rows are used)
- `.ocado_session/` — Playwright Chromium profile (do not commit)
- `logs/` — per-run logs
- `tests/` — pytest unit tests for the loaders (`pytest -v`)

## If something breaks

- Check the latest log in `logs/`.
- Most failures will be selector drift on Ocado's site. Edit the `SEL_*` constants at the top of `ocado_tuesday.py`.
```

- [ ] **Step 4: Run tests one more time**

```bash
source .venv/bin/activate
pytest -v
```

Expected: all 7 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add run_ocado.command README.md
git commit -m "feat: add macOS launcher and README"
```

---

## Task 9: Live smoke test

This is the only step that can't be automated — Ocado's real DOM has to be exercised by a human.

- [ ] **Step 1: Make sure `~/Downloads/ocado_tuesday.json` exists with a current `week`**

Create a small file with 1-2 items to keep the first test quick.

- [ ] **Step 2: Run `python ocado_tuesday.py`**

Expected first-time path:
- Chromium opens
- Script prints "Not logged in" — log in manually, press Enter
- Script clears reserved order, adds items, prints summary

- [ ] **Step 3: Fix any selector mismatches**

If a step times out, look at the log to see which selector. Edit the matching `SEL_*` constant in `ocado_tuesday.py` (look at the page in DevTools to find a stable selector). Re-run.

- [ ] **Step 4: Commit any selector fixes**

```bash
git add ocado_tuesday.py
git commit -m "fix: update SEL_* selectors after first live run"
```
