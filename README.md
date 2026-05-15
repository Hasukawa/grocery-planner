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

1. Generate `~/Code/grocery-planner/ocado_tuesday.json` from the checklist app. Format:
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
