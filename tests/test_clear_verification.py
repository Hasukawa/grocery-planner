"""Guard the clear-verification logic.

On 2026-08-11 a run logged "Verified: order is empty" while the order held
~65 items, then added 40 more on top. The old check read the header badge,
which is absent in the edit view, and treated absent as zero. These tests
pin the behaviour that must never regress: a full order is never reported
as empty, and "can't tell" is never reported as empty.
"""
import logging
import sys

from playwright.sync_api import sync_playwright

sys.path.insert(0, "/Users/grantastic/Code/grocery-planner")
from ocado_tuesday import count_order_items  # noqa: E402

log = logging.getLogger("test")

# The real edit view: no header badge at all, order lines exposed as
# "<name> click to edit" controls. This is the case that lied.
EDIT_VIEW_FULL = """
<div>
  <button>Grid view</button>
  <button>Empty trolley</button>
  <button>In your FavouritesM&S Organic Peppers click to edit</button>
  <button>In your FavouritesM&S Organic Cucumber click to edit</button>
  <button>In your FavouritesM&S Pink Lady Apples click to edit</button>
  <button>Check out to save changes</button>
</div>
"""

EDIT_VIEW_EMPTY = """
<div>
  <button>Grid view</button>
  <p>Your trolley is empty</p>
  <button>Check out to save changes</button>
</div>
"""

TROLLEY_DRAWER = """
<div>
  <div data-test="expanded-trolley-list-item"><a href="/products/a/1">A</a></div>
  <div data-test="expanded-trolley-list-item"><a href="/products/b/2">B</a></div>
</div>
"""

SUMMARY_TEXT = '<div><p>Summary</p><p>65 items</p><p>£311.42</p></div>'

HEADER_BADGE = '<div data-test="basket-counter"><span>53</span></div>'

# Nothing identifiable: must be indeterminate, NOT zero.
BLANK = '<div><p>Loading your order…</p></div>'

CASES = [
    ("edit view, 3 items (the case that lied)", EDIT_VIEW_FULL, 3, "edit-rows"),
    ("edit view, explicitly empty", EDIT_VIEW_EMPTY, 0, "empty-state-text"),
    ("trolley drawer, 2 rows", TROLLEY_DRAWER, 2, "trolley-rows"),
    ("summary panel text", SUMMARY_TEXT, 65, "summary-text"),
    ("header badge", HEADER_BADGE, 53, "header-badge"),
    ("nothing readable -> indeterminate", BLANK, None, "indeterminate"),
]


def main() -> int:
    failures = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        for label, html, want_count, want_how in CASES:
            page.set_content(html)
            count, how = count_order_items(page, log)
            ok = (count == want_count) and (how == want_how)
            print(f"  [{'ok ' if ok else 'FAIL'}] {label}: count={count} via={how}")
            if not ok:
                failures.append(f"{label}: want ({want_count}, {want_how}), got ({count}, {how})")
            # The cardinal rule: a non-empty order must never read as empty.
            if want_count not in (0, None) and count == 0:
                failures.append(f"{label}: DANGEROUS — full order reported as empty")
        browser.close()

    print()
    if failures:
        print("FAILURES:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print(f"PASS: all {len(CASES)} cases; full orders never read as empty, "
          "indeterminate stays indeterminate")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
