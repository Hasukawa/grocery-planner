"""Reproduce the qty>1 modal race, then prove add_locator_handler fixes it.

Mimics the real failure: clicking Add spawns the Favourites modal a beat
later — right as the script reaches for the quantity '+' — and the overlay
then intercepts every retry until the click times out.
"""
import logging
import sys

from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright

sys.path.insert(0, "/Users/grantastic/Code/grocery-planner")
from ocado_tuesday import install_modal_autodismiss  # noqa: E402

# Add spawns the blocking modal 800ms later; the '+' click starts before that.
HTML = """
<style>
  #portal-host .overlay {
    position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 999;
    display: flex; align-items: center; justify-content: center;
  }
  .dialog { background: #fff; padding: 20px; }
  button { padding: 10px; margin: 4px; }
</style>
<button id="add">Add</button>
<button id="plus" aria-label="Increase quantity of Test Item in trolley">+</button>
<div id="result">not clicked</div>
<div class="ReactModalPortal" id="portal-host"></div>
<script>
  document.getElementById('plus').addEventListener('click', () => {
    document.getElementById('result').textContent = 'PLUS CLICKED';
  });
  document.getElementById('add').addEventListener('click', () => {
    setTimeout(() => {
      document.getElementById('portal-host').innerHTML = `
        <div class="overlay">
          <div class="dialog">
            <h2>Favourites and Shopping Lists</h2>
            <button aria-label="Close modal" onclick="
              document.getElementById('portal-host').innerHTML=''">X</button>
          </div>
        </div>`;
    }, 200);
  });
</script>
"""


def attempt(page, log) -> tuple[bool, str]:
    page.set_content(HTML)
    page.click("#add")
    # The real sequence: the early dismiss found nothing, the modal landed
    # just after, and the '+' click then ran into it.
    page.wait_for_timeout(600)
    try:
        page.click("#plus", timeout=8_000)
    except PWTimeout as e:
        return False, str(e).splitlines()[0]
    return True, page.inner_text("#result")


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(message)s", stream=sys.stdout)
    log = logging.getLogger("test")
    failures = []

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)

        print("1. WITHOUT auto-dismiss (reproducing the bug)")
        page = browser.new_page()
        ok, detail = attempt(page, log)
        print(f"   click succeeded: {ok}  ({detail})")
        if ok:
            failures.append("expected the unguarded click to be blocked, but it passed "
                            "— the test no longer reproduces the bug")
        page.close()

        print("\n2. WITH auto-dismiss registered")
        page = browser.new_page()
        install_modal_autodismiss(page, log)
        ok, detail = attempt(page, log)
        print(f"   click succeeded: {ok}  ({detail})")
        if not ok:
            failures.append(f"auto-dismiss did not unblock the click: {detail}")
        elif detail != "PLUS CLICKED":
            failures.append(f"click landed on the wrong element: {detail!r}")
        page.close()

        print("\n3. Handler must NOT touch popups we need to answer")
        page = browser.new_page()
        install_modal_autodismiss(page, log)
        page.set_content("""
          <div class="ReactModalPortal">
            <div class="overlay" style="position:fixed;inset:0;background:#0008">
              <div class="dialog" style="background:#fff;padding:20px">
                <h2>Are you sure you want to empty your trolley?</h2>
                <button aria-label="Close modal">X</button>
                <button id="yes">Yes, empty trolley</button>
              </div>
            </div>
          </div>""")
        page.wait_for_timeout(1_200)
        still_there = page.locator('#yes').is_visible()
        print(f"   'Yes, empty trolley' still present: {still_there}")
        if not still_there:
            failures.append("handler wrongly dismissed the empty-trolley confirmation")
        page.close()
        browser.close()

    print()
    if failures:
        print("FAILURES:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("PASS: bug reproduced unguarded, fixed with handler, other popups untouched")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
