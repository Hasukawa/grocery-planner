"""One-off probe: figure out how to read the trolley reliably.

Answers two questions so we can preserve manually-added trolley items:
  1. Does Ocado expose the basket as JSON on the network? (gives us quantities)
  2. What DOM container holds ONLY the trolley items? (so we stop scraping
     the recommendation carousel by mistake)

Run with items already sitting in your trolley. Writes findings to
logs/probe-trolley.txt. Does not modify the basket.
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

ROOT = Path(__file__).parent
SESSION_DIR = ROOT / ".ocado_session"
LOGS_DIR = ROOT / "logs"
OUT_PATH = LOGS_DIR / "probe-trolley.txt"

SEL_BASKET_BTN = (
    '[data-test="basket-button"], [data-synthetics="basket-details-button"]'
)
# We know this button lives INSIDE the trolley panel — walking up from it
# finds the panel container without guessing at class names.
SEL_TROLLEY_CHECKOUT = (
    '[data-test="basket-checkout-button"], [data-synthetics="start-checkout-button"]'
)

INTERESTING_URL_BITS = ("basket", "trolley", "cart", "order")

lines: list[str] = []


def out(msg: str = "") -> None:
    print(msg)
    lines.append(msg)


def summarise_json(data, depth: int = 0, path: str = "") -> list[str]:
    """Describe the shape of a JSON blob, flagging anything that looks like line items."""
    found: list[str] = []
    if isinstance(data, dict):
        for k, v in data.items():
            here = f"{path}.{k}" if path else k
            if isinstance(v, list) and v and isinstance(v[0], dict):
                keys = sorted(v[0].keys())
                found.append(f"      LIST {here}  ({len(v)} entries)  item keys: {keys[:14]}")
                # Show one real entry so we can see if it has qty + product id
                sample = json.dumps(v[0], indent=2)[:600]
                found.append(f"        sample: {sample}")
            elif isinstance(v, (dict, list)) and depth < 3:
                found.extend(summarise_json(v, depth + 1, here))
    elif isinstance(data, list) and data and isinstance(data[0], dict) and depth < 3:
        found.extend(summarise_json(data[0], depth + 1, f"{path}[0]"))
    return found


def main() -> int:
    LOGS_DIR.mkdir(exist_ok=True)
    captured: list[tuple[str, object]] = []

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=str(SESSION_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = context.pages[0] if context.pages else context.new_page()

        def on_response(response):
            url = response.url
            if not any(bit in url.lower() for bit in INTERESTING_URL_BITS):
                return
            ctype = (response.headers or {}).get("content-type", "")
            if "json" not in ctype:
                return
            try:
                captured.append((url, response.json()))
            except Exception:  # noqa: BLE001
                pass

        page.on("response", on_response)

        out(f"Probe run {datetime.now():%Y-%m-%d %H:%M}")
        out("=" * 70)

        page.goto("https://www.ocado.com/", wait_until="domcontentloaded", timeout=30_000)
        try:
            page.wait_for_load_state("networkidle", timeout=20_000)
        except PWTimeout:
            pass

        # --- open the trolley panel -------------------------------------
        out("\n## Opening trolley panel")
        try:
            page.locator(SEL_BASKET_BTN).first.click(timeout=10_000)
            page.wait_for_timeout(3_000)
            out("   clicked basket button OK")
        except PWTimeout:
            out("   !! basket button not found — are you logged in?")

        # --- Q1: is there a JSON basket endpoint? -----------------------
        out("\n## Q1: network JSON responses mentioning basket/trolley/cart/order")
        if not captured:
            out("   (none captured)")
        for url, data in captured:
            out(f"\n   URL: {url[:150]}")
            hits = summarise_json(data)
            if hits:
                out("\n".join(hits))
            else:
                top = list(data.keys())[:20] if isinstance(data, dict) else type(data).__name__
                out(f"      (no obvious item list; top-level: {top})")

        # --- Q2: what container holds the trolley items? ----------------
        out("\n## Q2: DOM — walking up from the trolley's checkout button")
        info = page.evaluate(
            """() => {
            const btn = document.querySelector(
              '[data-test="basket-checkout-button"], [data-synthetics="start-checkout-button"]'
            );
            if (!btn) return { error: 'checkout button not found — trolley panel may not be open' };

            // Walk up the ancestors; at each level count product links found within.
            const rows = [];
            let el = btn;
            for (let i = 0; i < 8 && el; i++) {
                el = el.parentElement;
                if (!el) break;
                const links = el.querySelectorAll('a[href*="/products/"]');
                rows.push({
                    level: i + 1,
                    tag: el.tagName.toLowerCase(),
                    dataTest: el.getAttribute('data-test'),
                    dataSynthetics: el.getAttribute('data-synthetics'),
                    ariaLabel: el.getAttribute('aria-label'),
                    cls: (el.className || '').toString().slice(0, 80),
                    productLinks: links.length,
                });
            }

            // For comparison: how many product links exist on the WHOLE page?
            const pageWide = document.querySelectorAll('a[href*="/products/"]').length;

            // Look for anything that self-identifies as a trolley/basket item row.
            const candidates = [...document.querySelectorAll('[data-test], [data-synthetics]')]
                .map(e => e.getAttribute('data-test') || e.getAttribute('data-synthetics'))
                .filter(v => v && /basket|trolley|cart/i.test(v));
            const counts = {};
            for (const c of candidates) counts[c] = (counts[c] || 0) + 1;

            return { rows, pageWide, attrCounts: counts };
        }"""
        )

        if info.get("error"):
            out(f"   !! {info['error']}")
        else:
            out(f"   product links on WHOLE page: {info['pageWide']}"
                "   <- what the old buggy scrape counted")
            out("\n   ancestors of the checkout button (find the level whose"
                " link count == your real trolley item count):")
            for r in info["rows"]:
                out(f"     L{r['level']}: <{r['tag']}> links={r['productLinks']:<4}"
                    f" data-test={r['dataTest']!r} data-synthetics={r['dataSynthetics']!r}"
                    f" class={r['cls']!r}")
            out("\n   attributes mentioning basket/trolley/cart (name: count):")
            for k, v in sorted(info["attrCounts"].items()):
                out(f"     {k}: {v}")

        out("\n" + "=" * 70)
        out("Leaving browser open. Close the window when done.")
        OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
        print(f"\nWrote {OUT_PATH}")

        try:
            page.wait_for_event("close", timeout=0)
        except Exception:  # noqa: BLE001
            pass
        try:
            context.close()
        except Exception:  # noqa: BLE001
            pass

    OUT_PATH.write_text("\n".join(lines), encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
