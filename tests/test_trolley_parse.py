"""Verify JS_READ_TROLLEY against a mock of Ocado's trolley drawer markup.

Mirrors the real structure found by probe_trolley.py:
  - rows tagged data-test="expanded-trolley-list-item"
  - two product links per row (concealed image link + title link)
  - quantity in data-test="quantity-in-basket"
  - a recommendation carousel OUTSIDE the rows, which must be ignored
"""
import re
import sys
from playwright.sync_api import sync_playwright

sys.path.insert(0, "/Users/grantastic/Code/grocery-planner")
from ocado_tuesday import JS_READ_TROLLEY  # noqa: E402

HTML = """
<div data-test="expanded-trolley-drawer">
  <div data-test="expanded-trolley-list-item">
    <a href="/products/leerdammer-original/27775011" tabindex="-1" aria-hidden="true"></a>
    <a href="/products/leerdammer-original/27775011">Leerdammer Original Slices</a>
    <span data-test="quantity-in-basket">2</span>
  </div>
  <div data-test="expanded-trolley-list-item">
    <a href="/products/ocado-cauliflower/65350011" tabindex="-1" aria-hidden="true"></a>
    <a href="/products/ocado-cauliflower/65350011">Ocado Cauliflower</a>
    <span data-test="quantity-in-basket">1</span>
  </div>
  <!-- quantity rendered with surrounding text, and a query string on the href -->
  <div data-test="expanded-trolley-list-item">
    <a href="/products/gails-seeded-crackers/53550011?from=trolley" tabindex="-1"></a>
    <a href="/products/gails-seeded-crackers/53550011?from=trolley">GAIL's Seeded Crackers</a>
    <span data-test="quantity-in-basket">Qty: 3</span>
  </div>

  <!-- The trap that broke the old scrape: recommendations inside the drawer -->
  <section class="recommendations">
    <a href="/products/random-upsell-one/11111011">Buy this too</a>
    <a href="/products/random-upsell-two/22222011">And this</a>
    <a href="/products/random-upsell-three/33333011">And this as well</a>
  </section>
</div>
"""


def parse(entry):
    """Mirror the Python-side parsing in capture_trolley_items()."""
    digits = re.search(r"\d+", entry.get("qtyRaw") or "")
    return int(digits.group()) if digits else 1


def main() -> int:
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.set_content(HTML)

        page_wide = page.locator('a[href*="/products/"]').count()
        rows = page.evaluate(JS_READ_TROLLEY)
        browser.close()

    print(f"product links page-wide (what the OLD code scraped): {page_wide}")
    print(f"rows returned by JS_READ_TROLLEY:                    {len(rows)}")
    print()

    failures = []
    if page_wide != 9:
        failures.append(f"expected 9 page-wide links in fixture, got {page_wide}")
    if len(rows) != 3:
        failures.append(f"expected 3 trolley rows, got {len(rows)}")

    expected = [
        ("Leerdammer Original Slices", "https://www.ocado.com/products/leerdammer-original/27775011", 2),
        ("Ocado Cauliflower", "https://www.ocado.com/products/ocado-cauliflower/65350011", 1),
        ("GAIL's Seeded Crackers", "https://www.ocado.com/products/gails-seeded-crackers/53550011", 3),
    ]
    for i, (want_name, want_url, want_qty) in enumerate(expected):
        if i >= len(rows):
            failures.append(f"row {i}: missing")
            continue
        got = rows[i]
        qty = parse(got)
        print(f"  row {i}: {qty} x {got['name']!r}")
        print(f"          {got['href']}")
        if got["name"] != want_name:
            failures.append(f"row {i} name: want {want_name!r}, got {got['name']!r}")
        if got["href"] != want_url:
            failures.append(f"row {i} url: want {want_url!r}, got {got['href']!r}")
        if qty != want_qty:
            failures.append(f"row {i} qty: want {want_qty}, got {qty} (raw {got['qtyRaw']!r})")

    urls = " ".join(r["href"] for r in rows)
    if "upsell" in urls:
        failures.append("LEAK: a recommendation link was captured as a trolley item")

    print()
    if failures:
        print("FAILURES:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("PASS: 3 rows, correct names/urls/quantities, 6 carousel links ignored")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
