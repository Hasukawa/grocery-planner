"""Ocado Tuesday automation script."""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Literal

import openpyxl
from playwright.sync_api import Page, TimeoutError as PWTimeout, sync_playwright


ROOT = Path(__file__).parent
XLSX_PATH = ROOT / "Ocado_Order_Manager.xlsx"
JSON_PATH = ROOT / "ocado_tuesday.json"
JSON_DOWNLOADS_FALLBACK = Path.home() / "Downloads" / "ocado_tuesday.json"
SESSION_DIR = ROOT / ".ocado_session"
LOGS_DIR = ROOT / "logs"

URL_HOME = "https://www.ocado.com/"
# Use the home page as the anchor for login checks + as a known-good page to search from.
URL_RESERVED = URL_HOME
# Modern orders flow:
#   /orders                       → list of upcoming Wednesday cards
#   /orders/{ID}/details          → click into a specific week
#   then click "Edit order"       → enters editable basket
URL_ORDERS_LIST = "https://www.ocado.com/orders"
LOGIN_URL_FRAGMENTS = ("/login", "/signin", "sign-in", "log-in", "accounts.ocado", "/auth")
SKIP_CLEAR = True  # Set False once we've smoke-tested the clear-basket flow

# Selectors — comma-separated fallbacks. Tweak on first run if Ocado has changed.
SEL_SEARCH_INPUT = (
    'input[placeholder*="Find a product" i], '
    'input[placeholder*="Find" i], '
    'input[aria-label*="search" i], '
    'input[role="searchbox"], '
    'input[role="combobox"], '
    '[role="search"] input, '
    'input[type="search"], '
    'input[name="search"], '
    '[data-testid="search-input"]'
)
SEL_SEARCH_SUBMIT = 'button[type="submit"][aria-label*="earch"], button:has-text("Search")'
SEL_PRODUCT_CARD = (
    'a[href*="/products/"], '
    '[data-testid*="product-tile"] a, '
    '[data-testid*="product-card"] a, '
    'article[data-product-id] a, '
    'li[data-product-id] a, '
    '[data-sku] a'
)
SEL_PRODUCT_TITLE = 'h1, [data-testid="product-title"]'
SEL_ADD_BUTTON = 'button:has-text("Add"), button[data-testid*="add-to-trolley"]'
SEL_QTY_PLUS = 'button[aria-label*="ncrease"], button:has-text("+")'
SEL_REMOVE_ALL = 'button:has-text("Remove all"), button:has-text("Empty trolley"), button:has-text("Clear trolley")'
SEL_REMOVE_CONFIRM = 'button:has-text("Yes"), button:has-text("Confirm"), button:has-text("Remove")'
# Positive logged-in markers — only appear when authenticated
SEL_LOGGED_IN_MARKER = 'a:has-text("Sign out"), a:has-text("Log out"), button:has-text("Sign out"), button:has-text("Log out")'
# Logged-out / login-page markers — fields and buttons only present when NOT authenticated
SEL_LOGGED_OUT_MARKER = (
    'input[type="password"], '
    'a:has-text("Sign in"), a:has-text("Log in"), '
    'button:has-text("Sign in"), button:has-text("Log in")'
)

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


def resolve_tier2_json_path(log: logging.Logger) -> Path | None:
    """Find the Tier 2 JSON, preferring the project root over ~/Downloads.

    If the file is only present in ~/Downloads, move it to the project root so
    future runs find it there directly. Returns None if neither location has it.
    """
    if JSON_PATH.exists():
        return JSON_PATH
    if JSON_DOWNLOADS_FALLBACK.exists():
        log.info("Found %s in Downloads — moving to project root", JSON_DOWNLOADS_FALLBACK.name)
        JSON_DOWNLOADS_FALLBACK.rename(JSON_PATH)
        return JSON_PATH
    return None


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


def login_state(page: Page) -> Literal["in", "out", "unknown"]:
    """Positive-marker login detection. Returns 'in' / 'out' / 'unknown'."""
    url = page.url
    if any(frag in url for frag in LOGIN_URL_FRAGMENTS):
        return "out"
    try:
        if page.locator(SEL_LOGGED_IN_MARKER).first.is_visible(timeout=1_500):
            return "in"
    except PWTimeout:
        pass
    try:
        if page.locator(SEL_LOGGED_OUT_MARKER).first.is_visible(timeout=1_500):
            return "out"
    except PWTimeout:
        pass
    return "unknown"


def ensure_logged_in(page: Page) -> None:
    log = logging.getLogger("ocado")
    page.goto(URL_RESERVED, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
    try:
        page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    except PWTimeout:
        pass
    for attempt in (1, 2):
        state = login_state(page)
        log.info("Login check: state=%s url=%s", state, page.url)
        if state == "in":
            return
        if state == "unknown":
            log.warning("Login state unclear — please confirm manually.")
        else:
            log.warning("Not logged in (attempt %d).", attempt)
        print(f">>> Browser is at: {page.url}")
        print(">>> If not already logged in, log in now. Then press Enter here to continue.")
        input(">>> ")
        page.goto(URL_RESERVED, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
        try:
            page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
        except PWTimeout:
            pass
    if login_state(page) == "out":
        raise RuntimeError("Still not logged in after retry — aborting.")


def clear_reserved_order(page: Page, target_date: date) -> None:
    """Navigate /orders → matching Wednesday card → Edit order → remove all.

    Ocado lists upcoming orders as cards with text like 'Wed 20 May 9:00pm - 10:00pm'.
    We match on 'Wed {DD} {Mon}' built from target_date.
    """
    log = logging.getLogger("ocado")
    log.info("Clearing reserved order for %s…", target_date.isoformat())
    page.goto(URL_ORDERS_LIST, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
    try:
        page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    except PWTimeout:
        pass

    day_str = f"Wed {target_date.day} {target_date.strftime('%b')}"
    log.info("Looking for order card: '%s'", day_str)
    card = page.locator(f':has-text("{day_str}")').last  # `.last` skips the page heading
    try:
        card.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
    except PWTimeout:
        log.warning("No order card found for '%s' — skipping clear.", day_str)
        return
    card.click()
    try:
        page.wait_for_url("**/orders/*/details*", timeout=NAVIGATION_TIMEOUT_MS)
    except PWTimeout:
        log.warning("Did not reach /orders/*/details after click — skipping clear.")
        return

    edit_btn = page.locator('button:has-text("Edit order"), a:has-text("Edit order")').first
    try:
        edit_btn.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        edit_btn.click()
    except PWTimeout:
        log.warning("No 'Edit order' button — basket may already be in edit mode or empty.")
    try:
        page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    except PWTimeout:
        pass

    remove_btn = page.locator(SEL_REMOVE_ALL).first
    try:
        remove_btn.wait_for(state="visible", timeout=5_000)
        remove_btn.click()
        log.info("Clicked 'Remove all'.")
        confirm = page.locator(SEL_REMOVE_CONFIRM).first
        try:
            confirm.wait_for(state="visible", timeout=3_000)
            confirm.click()
            log.info("Confirmed removal.")
        except PWTimeout:
            log.info("No confirmation dialog — that's fine.")
    except PWTimeout:
        log.info("No 'Remove all' button found in edit view — basket may already be empty.")


@dataclass
class Result:
    item: Item
    status: Literal["ok", "not_found", "error"]
    detail: str = ""
    product_url: str = ""


# Matches sizes that confuse Ocado's search: '300g', '1.5l', '8x330ml', '6 x 90g', etc.
# Ocado returns zero results when the search term contains a specific size.
_SIZE_PATTERN = re.compile(
    r'\s*\b\d+(?:\.\d+)?\s*(?:[x×]\s*\d+(?:\.\d+)?\s*)?(?:g|kg|mg|ml|cl|l|oz|lb|pints?|pt)\b\s*',
    re.IGNORECASE,
)


def strip_size_suffix(term: str) -> str:
    """Remove '300g', '1kg', '750ml', '2 pints' etc. from a search term."""
    cleaned = _SIZE_PATTERN.sub(' ', term)
    return re.sub(r'\s+', ' ', cleaned).strip()


SEL_MODAL_CLOSE = (
    'button[aria-label="Close" i], '
    'button[aria-label*="close" i], '
    '[role="dialog"] button:has-text("Close"), '
    '[role="dialog"] button:has-text("×"), '
    '[role="dialog"] button:has-text("Not now"), '
    '[role="dialog"] button:has-text("No thanks")'
)


def dismiss_modals(page: Page, log: logging.Logger) -> None:
    """Close any open modal/dialog overlay. Safe to call when none is present."""
    # Press Escape first — handles most React modals
    try:
        if page.locator('.ReactModalPortal, [role="dialog"]').first.is_visible(timeout=500):
            page.keyboard.press("Escape")
            log.info("   pressed Escape to dismiss modal")
    except PWTimeout:
        pass
    # If still present, try clicking a close button
    try:
        close = page.locator(SEL_MODAL_CLOSE).first
        if close.is_visible(timeout=500):
            close.click()
            log.info("   clicked modal close button")
    except PWTimeout:
        pass


def _try_search_suggestion(page: Page, log: logging.Logger) -> bool:
    """If page shows 'No results for', click the first suggestion chip and return True."""
    no_results = page.get_by_text("No results for", exact=False).first
    try:
        no_results.wait_for(state="visible", timeout=2_000)
    except PWTimeout:
        return False
    suggestion = no_results.locator('xpath=following::button[1]')
    try:
        suggestion.wait_for(state="visible", timeout=2_000)
    except PWTimeout:
        return False
    text = (suggestion.text_content() or "").strip()
    log.info("   no results — clicking suggestion: %r", text)
    suggestion.click()
    try:
        page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
    except PWTimeout:
        pass
    return True


def _titles_roughly_match(actual_title: str, search_term: str) -> bool:
    """True if the longest word in search_term appears (case-insensitive) in actual_title."""
    words = [w for w in search_term.split() if len(w) >= 4]
    if not words:
        return True
    longest = max(words, key=len)
    return longest.lower() in actual_title.lower()


def add_item(page: Page, item: Item, capture_only: bool = False) -> Result:
    """Search for or navigate to the item's product page and add it to the basket.

    If capture_only=True, navigate to the product page but do NOT add to basket
    or change quantity. Used to populate URL cache without touching the basket.
    """
    log = logging.getLogger("ocado")
    log.info("→ %s [%s] qty=%d%s", item.name, item.source, item.qty, " (capture-only)" if capture_only else "")
    dismiss_modals(page, log)
    try:
        if item.notes.startswith("https://www.ocado.com"):
            page.goto(item.notes, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
        else:
            search_term = strip_size_suffix(item.search_term)
            if search_term != item.search_term:
                log.info("   search term cleaned: '%s' → '%s'", item.search_term, search_term)
            search_box = page.locator(SEL_SEARCH_INPUT).first
            search_box.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            search_box.fill(search_term)
            search_box.press("Enter")
            page.wait_for_load_state("networkidle", timeout=NAVIGATION_TIMEOUT_MS)
            first_card = page.locator(SEL_PRODUCT_CARD).first
            try:
                first_card.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            except PWTimeout:
                # "No results" fallback: Ocado often shows a shorter suggestion chip
                # ("Here are some product recommendations…"). Try clicking the first
                # such chip and re-check for products.
                if _try_search_suggestion(page, log):
                    try:
                        first_card.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
                    except PWTimeout:
                        return Result(item, "not_found", f"no results for '{search_term}' (suggestion also empty)")
                else:
                    return Result(item, "not_found", f"no results for '{search_term}'")
            first_card.click()

        title_loc = page.locator(SEL_PRODUCT_TITLE).first
        title_loc.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        actual = (title_loc.text_content() or "").strip()
        if not _titles_roughly_match(actual, item.search_term):
            log.warning("Title mismatch: wanted ~'%s', got '%s' — continuing", item.search_term, actual)

        if capture_only:
            log.info("   captured URL: %s", page.url)
            return Result(item, "ok", actual, product_url=page.url)

        add_btn = page.locator(SEL_ADD_BUTTON).first
        add_btn.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
        add_btn.click()
        log.info("   added 1")

        for n in range(item.qty - 1):
            plus = page.locator(SEL_QTY_PLUS).first
            plus.wait_for(state="visible", timeout=ELEMENT_TIMEOUT_MS)
            plus.click()
            log.info("   qty +1 (now %d)", n + 2)

        return Result(item, "ok", actual, product_url=page.url)
    except PWTimeout as e:
        return Result(item, "error", f"timeout: {e}")
    except Exception as e:  # noqa: BLE001
        return Result(item, "error", f"{type(e).__name__}: {e}")


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


def write_product_urls(xlsx_path: Path, results: list[Result], log: logging.Logger) -> int:
    """Write captured product URLs back to the spreadsheet.

    - Tier 1 results → "Tier 1 – Essentials" sheet, Notes column (col 5)
    - Tier 2 results → "Rotation Pool" sheet, Notes column (col 7)

    Matches by exact Item Name. Always overwrites existing Notes content for
    matching rows. Returns the number of cells updated across both sheets.
    """
    tier1_urls = {
        r.item.name: r.product_url
        for r in results
        if r.item.source == "tier1" and r.status == "ok" and r.product_url
    }
    tier2_urls = {
        r.item.name: r.product_url
        for r in results
        if r.item.source == "tier2" and r.status == "ok" and r.product_url
    }
    if not tier1_urls and not tier2_urls:
        log.info("No URLs to write back.")
        return 0
    wb = openpyxl.load_workbook(xlsx_path)
    updated = 0
    if tier1_urls:
        ws = wb["Tier 1 – Essentials"]
        # Notes is column 5 (index 4)
        for row in ws.iter_rows(min_row=3):
            name_cell = row[0]
            if not name_cell.value:
                continue
            name = str(name_cell.value).strip()
            if name in tier1_urls:
                row[4].value = tier1_urls[name]
                updated += 1
    if tier2_urls:
        ws = wb["Rotation Pool"]
        # Notes is column 7 (index 6)
        matched = set()
        for row in ws.iter_rows(min_row=3):
            name_cell = row[0]
            if not name_cell.value:
                continue
            name = str(name_cell.value).strip()
            if name in tier2_urls:
                row[6].value = tier2_urls[name]
                matched.add(name)
                updated += 1
        unmatched = set(tier2_urls) - matched
        if unmatched:
            log.warning(
                "Tier 2 names not found in Rotation Pool sheet (URL not written): %s",
                ", ".join(sorted(unmatched)),
            )
    wb.save(xlsx_path)
    return updated


def confirm_write_urls(count: int) -> bool:
    print(f"\n{count} product URL(s) captured this run.")
    answer = input("Write them back to the spreadsheet's Notes column? [y/N]: ").strip().lower()
    return answer == "y"


def confirm_stale_week(week_str: str, expected: date) -> bool:
    print(f"\n⚠  Tier 2 JSON week is '{week_str}', expected '{expected.isoformat()}'.")
    answer = input("   Proceed anyway? [y/N]: ").strip().lower()
    return answer == "y"


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Ocado Tuesday automation")
    parser.add_argument(
        "--capture-only",
        action="store_true",
        help="Visit each Tier 1 item's product page to grab its URL and write it back to the spreadsheet's Notes column. Skips clearing the basket and skips adding items. Items that already have a URL in Notes are skipped.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    log_path = setup_logging()
    log = logging.getLogger("ocado")
    log.info("Run starting. Log: %s", log_path)
    if args.capture_only:
        log.warning("CAPTURE-ONLY mode: visiting items to grab URLs. Basket will NOT be touched.")

    tier1 = load_tier1(XLSX_PATH)
    log.info("Loaded %d Tier 1 items", len(tier1))

    if args.capture_only:
        # In capture mode we only need Tier 1; skip items that already have a URL.
        items_to_visit = [
            it for it in tier1
            if not it.notes.startswith("https://www.ocado.com")
        ]
        skipped = len(tier1) - len(items_to_visit)
        log.info("Capturing URLs for %d items (%d already have a URL, skipped)", len(items_to_visit), skipped)
        expected_week = upcoming_wednesday(date.today())
    else:
        json_path = resolve_tier2_json_path(log)
        if json_path is None:
            log.error("Tier 2 JSON not found at %s or %s — aborting.", JSON_PATH, JSON_DOWNLOADS_FALLBACK)
            return 1
        tier2, week_str = load_tier2(json_path)
        log.info("Loaded %d Tier 2 items (week=%s) from %s", len(tier2), week_str, json_path.name)

        expected_week = upcoming_wednesday(date.today())
        if week_str != expected_week.isoformat():
            if not confirm_stale_week(week_str, expected_week):
                log.info("User declined stale JSON — aborting.")
                return 1
        items_to_visit = tier1 + tier2

    all_items = items_to_visit
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
            if args.capture_only:
                log.info("Skipping basket clear (capture-only mode).")
                page.goto(URL_HOME, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
            elif SKIP_CLEAR:
                log.warning("SKIP_CLEAR=True — basket-clearing disabled. Clear the basket manually in the browser, then press Enter.")
                input(">>> Press Enter once the basket is empty (or to proceed anyway): ")
                page.goto(URL_HOME, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT_MS)
            else:
                clear_reserved_order(page, expected_week)
            for i, item in enumerate(all_items, 1):
                result = add_item(page, item, capture_only=args.capture_only)
                results.append(result)
                if result.status == "ok":
                    log.info("[%d/%d] OK: %s", i, len(all_items), item.name)
                elif result.status == "not_found":
                    log.warning("[%d/%d] NOT FOUND: %s (%s)", i, len(all_items), item.name, result.detail)
                else:
                    log.error("[%d/%d] ERROR: %s (%s)", i, len(all_items), item.name, result.detail)
        except Exception as e:  # noqa: BLE001
            log.exception("Run aborted: %s", e)

        print_summary(results, log_path)

        capturable = sum(1 for r in results if r.status == "ok" and r.product_url)
        if capturable > 0 and confirm_write_urls(capturable):
            try:
                n = write_product_urls(XLSX_PATH, results, log)
                log.info("Wrote %d URLs to %s", n, XLSX_PATH.name)
            except PermissionError:
                log.error("Could not write to %s — is Excel/Numbers holding it open? Close it and try again.", XLSX_PATH.name)
            except Exception as e:  # noqa: BLE001
                log.exception("Failed to write URLs: %s", e)

        print("\nBrowser left open for review. Close the Chromium window when done — your session will be saved.")
        try:
            page.wait_for_event("close", timeout=0)
        except Exception:  # noqa: BLE001
            pass
        try:
            context.close()
        except Exception:  # noqa: BLE001
            pass

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
