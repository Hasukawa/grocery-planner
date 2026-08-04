# Ocado Tuesday Automation — Design

**Date:** 2026-05-14
**Status:** Approved

## Purpose

Weekly Tuesday-evening script that fills the user's Ocado basket with two sets of items:
- **Tier 1**: auto-add every week (from `Ocado_Order_Manager.xlsx`, sheet "Tier 1 – Essentials", rows where `Active? = "Yes"`)
- **Tier 2**: this week's YES choices (from `~/Downloads/ocado_tuesday.json`)

User double-clicks a launcher, Playwright drives Chromium, items land in basket, user reviews and checks out manually.

## Deliverables

- `ocado_tuesday.py` — main script
- `run_ocado.command` — macOS double-clickable launcher
- `requirements.txt` — `playwright`, `openpyxl`
- `README.md` — first-run setup
- `.ocado_session/` — Playwright persistent profile (auto-created)
- `logs/` — per-run logs (auto-created)

## File layout

```
~/Code/grocery-planner/
├── ocado_tuesday.py
├── run_ocado.command
├── requirements.txt
├── README.md
├── Ocado_Order_Manager.xlsx        (existing)
├── .ocado_session/                 (Playwright user data dir)
└── logs/
    └── run-YYYY-MM-DD-HHMM.log
```

## Script structure (`ocado_tuesday.py`)

Single file, top-to-bottom readable. Section order:

1. **Imports + constants** — selectors, URLs, paths, timeouts, all at the top
2. **`Item` dataclass** — `name, qty, search_term, notes, source` (`"tier1" | "tier2"`)
3. **Loaders** — `load_tier1(xlsx_path) -> list[Item]`, `load_tier2(json_path) -> list[Item]`
4. **Browser helpers** — `ensure_logged_in(page)`, `clear_reserved_order(page)`, `add_item(page, item) -> Result`
5. **Main** — load data, launch persistent context, run flow, print summary, wait for Enter to close

Selectors live as named constants at the top so first-run tweaks are one-file changes:
```python
SEL_SEARCH_INPUT = '[data-testid="search-input"], input[name="search"]'
SEL_REMOVE_ALL   = 'button:has-text("Remove all"), button:has-text("Empty trolley")'
SEL_PRODUCT_CARD = '[data-testid^="product-tile"]'
```

Use comma-separated fallback selectors where guessing — Playwright tries each in order.

## Per-item flow (`add_item`)

1. If `item.notes` starts with `https://www.ocado.com` → `page.goto(notes_url)`
2. Else → fill `SEL_SEARCH_INPUT` with `item.search_term`, press Enter, wait for results, click first product card
3. On product page: wait for product title visible. Fuzzy check title vaguely matches `search_term` (lowercase substring of one main word) — if not, log warning but continue
4. Click "Add" once; click "+" `(qty - 1)` more times
5. Return one of: `Ok`, `NotFound`, `Error(msg)`

Each item is wrapped in its own try/except so one bad item doesn't kill the run.

## Login & session handling

- `launch_persistent_context(user_data_dir=".ocado_session")` keeps cookies/localStorage across runs
- **First run**: log in manually once; session persists
- **`ensure_logged_in(page)`**: navigate to `/webshop/reservedOrder.do`, check URL for `/login`/`/signin` or visible "Sign in" button. If not authed: print `>>> Not logged in. Log in in the browser, then press Enter to continue...` and `input()`. Re-check after Enter; one retry; then bail with clear error
- Same check is invoked again if a mid-run page action times out (treat as possible session expiry)

## Reserved-order clearing

`clear_reserved_order(page)`:
1. Goto `https://www.ocado.com/webshop/reservedOrder.do`
2. Try `SEL_REMOVE_ALL` (multiple fallback texts)
3. If found → click, handle confirmation dialog if it appears, wait for empty-basket state
4. If not found → assume already empty, log it, continue

## Tier 2 JSON handling

- File missing → exit with clear error
- File present but `"week"` field doesn't match the upcoming Wednesday's date → print loud warning, prompt y/n to proceed
- File present and current → use as-is

## Run summary

Printed before "press Enter to close":

```
================ SUMMARY ================
Tier 1: 18/20 added
Tier 2:  6/7  added

Not found (1):
  - M&S Some Obscure Thing [tier1]    search: "..."

Errors (2):
  - M&S Raspberries [tier1]: timeout adding to basket
  - Whatever [tier2]: title mismatch (got "...")

Full log: logs/run-2026-05-19-2103.log
=========================================

Browser left open for review. Press Enter to close.
```

## Launcher (`run_ocado.command`)

```bash
#!/bin/bash
cd "$(dirname "$0")"
source .venv/bin/activate 2>/dev/null || {
  echo "Virtualenv missing — run setup steps from README.md first."
  read -n1
  exit 1
}
python ocado_tuesday.py
echo
read -p "Done. Press Enter to close this window."
```

`chmod +x` so it's double-clickable from Finder.

## README first-run setup

1. `cd ~/Code/grocery-planner`
2. `python3 -m venv .venv && source .venv/bin/activate`
3. `pip install -r requirements.txt`
4. `playwright install chromium`
5. First run: `python ocado_tuesday.py` — browser opens, log in to Ocado manually, press Enter in terminal, script proceeds. Session saved
6. Subsequent runs: double-click `run_ocado.command`

## Robustness

- Each `add_item` in its own try/except → record status, continue
- Playwright `expect()` / `wait_for_selector` everywhere — no `time.sleep` except where no good signal exists
- Default timeouts: 15s element waits, 30s navigations
- Log every step (item name, selector tried, outcome) to stdout and `logs/run-*.log`
- Global try/except around main → on any unexpected crash, still print summary and log path before re-raising

## Out of scope

- No checkout — items land in basket; user checks out manually
- No Tier 3 / Rotation Pool — only "Tier 1 – Essentials" sheet + JSON
- No de-dup against existing basket — we clear first, then add fresh
- No price/budget tracking
- No quantity-vs-pack-size validation
