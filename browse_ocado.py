"""Open a Chromium window with the saved Ocado session for manual browsing."""
from pathlib import Path
from playwright.sync_api import sync_playwright

ROOT = Path(__file__).parent
SESSION_DIR = ROOT / ".ocado_session"

with sync_playwright() as p:
    context = p.chromium.launch_persistent_context(
        user_data_dir=str(SESSION_DIR),
        headless=False,
        viewport={"width": 1400, "height": 900},
    )
    page = context.pages[0] if context.pages else context.new_page()
    page.goto("https://www.ocado.com/")
    page.wait_for_event("close", timeout=0)
