"""Verify the quantity read handles every markup shape we might hit.

The live probe showed textContent empty for all 19 rows, so the reader now
tries value-prop, nested input, aria-valuenow and attributes. Each is
exercised here.
"""
import re
import sys
from playwright.sync_api import sync_playwright

sys.path.insert(0, "/Users/grantastic/Code/grocery-planner")
from ocado_tuesday import JS_READ_TROLLEY  # noqa: E402

P = "https://www.ocado.com/products/x/1"

CASES = [
    ("input value (most likely real shape)",
     '<input data-test="quantity-in-basket" value="4">', 4, "value-prop"),
    ("nested input inside wrapper",
     '<div data-test="quantity-in-basket"><input value="7"></div>', 7, "nested-input-value"),
    ("aria-valuenow on nested element",
     '<div data-test="quantity-in-basket"><span aria-valuenow="5"></span></div>', 5,
     "nested-aria-valuenow"),
    ("plain text",
     '<span data-test="quantity-in-basket">3</span>', 3, "textContent"),
    ("attribute only",
     '<div data-test="quantity-in-basket" aria-valuenow="6"></div>', 6, "attr:aria-valuenow"),
    ("row counter input fallback (no quantity-in-basket at all)",
     '<input data-test="counter:input" value="9">', 9, "row-counter-input"),
    ("genuinely empty -> default 1, flagged",
     '<span data-test="quantity-in-basket"></span>', 1, "EMPTY"),
]


def build(markup: str) -> str:
    return f"""
    <div data-test="expanded-trolley-list-item">
      <a href="/products/x/1" tabindex="-1"></a>
      <a href="/products/x/1">Test Product</a>
      {markup}
    </div>"""


def main() -> int:
    failures = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        for label, markup, want_qty, want_how in CASES:
            page.set_content(build(markup))
            rows = page.evaluate(JS_READ_TROLLEY)
            if len(rows) != 1:
                failures.append(f"{label}: expected 1 row, got {len(rows)}")
                continue
            got = rows[0]
            digits = re.search(r"\d+", got.get("qtyRaw") or "")
            qty = int(digits.group()) if digits else 1
            flagged = "" if digits else "  (flagged as unreadable)"
            status = "ok " if (qty == want_qty and got["qtyHow"] == want_how) else "FAIL"
            print(f"  [{status}] {label}")
            print(f"         qty={qty} via={got['qtyHow']}{flagged}")
            if qty != want_qty:
                failures.append(f"{label}: want qty {want_qty}, got {qty}")
            if got["qtyHow"] != want_how:
                failures.append(f"{label}: want how={want_how}, got {got['qtyHow']}")
        browser.close()

    print()
    if failures:
        print("FAILURES:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print(f"PASS: all {len(CASES)} quantity markup shapes handled")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
