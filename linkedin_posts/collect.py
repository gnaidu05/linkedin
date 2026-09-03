"""Live collection layer: drive a browser the human logs into by hand.

This is the ONE part of the pipeline that talks to LinkedIn. It is kept
deliberately thin -- its whole job is to return the HTML of a search-results
page -- so that everything downstream (parsing, storage, Excel) can be tested
without a network or a browser.

Important, by design:
  * The browser opens visibly (headed). The human logs in themselves.
  * This code never reads, types, or stores a username or password. LinkedIn's
    own session cookies live in the browser profile directory, created when the
    person logs in by hand -- the same as a normal browser profile.
  * Automating LinkedIn's interface is against LinkedIn's Terms of Service and
    can get an account restricted. See the README.
"""

from __future__ import annotations

import time
from urllib.parse import quote_plus

from .config import Config
from .parse import POST_CONTAINER_SELECTORS

SEARCH_URL = "https://www.linkedin.com/search/results/content/?keywords={term}"
FEED_URL = "https://www.linkedin.com/feed/"


def _is_logged_in(page) -> bool:
    """Heuristic: logged-in sessions can reach the feed without a login form."""
    url = page.url
    if "/login" in url or "/uas/login" in url or "/checkpoint" in url:
        return False
    # A visible password field means we're still on an auth screen.
    try:
        if page.locator("input[type='password']").count() > 0:
            return False
    except Exception:
        pass
    return True


def collect_html(search_term: str, config: Config) -> str:
    """Open a browser, let the human log in, run the search, return page HTML.

    Raises RuntimeError if the human does not finish logging in within the
    configured timeout.
    """
    # Imported lazily so the rest of the package works without Playwright.
    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            user_data_dir=config.user_data_dir,
            headless=False,
        )
        try:
            page = context.pages[0] if context.pages else context.new_page()

            page.goto(FEED_URL, wait_until="domcontentloaded")
            if not _is_logged_in(page):
                print(
                    "\nA browser window has opened. Please log in to LinkedIn "
                    "there.\nWaiting for you to finish "
                    f"(up to {config.login_timeout_seconds}s)..."
                )
                _wait_for_login(page, config.login_timeout_seconds)
                print("Login detected. Continuing.\n")

            page.goto(
                SEARCH_URL.format(term=quote_plus(search_term)),
                wait_until="domcontentloaded",
            )
            _scroll_until_enough(page, config)
            return page.content()
        finally:
            context.close()


def _wait_for_login(page, timeout_seconds: int) -> None:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        if _is_logged_in(page):
            return
        time.sleep(1.0)
    raise RuntimeError(
        "Timed out waiting for login. Re-run and complete login in the browser."
    )


def _count_posts(page) -> int:
    selector = ", ".join(POST_CONTAINER_SELECTORS)
    try:
        return page.locator(selector).count()
    except Exception:
        return 0


def _scroll_until_enough(page, config: Config) -> None:
    """Scroll to trigger lazy loading until we have enough posts or stall."""
    stagnant_rounds = 0
    last_count = _count_posts(page)
    while last_count < config.max_results and stagnant_rounds < 3:
        page.mouse.wheel(0, 20000)
        time.sleep(config.scroll_pause_seconds)
        count = _count_posts(page)
        if count <= last_count:
            stagnant_rounds += 1
        else:
            stagnant_rounds = 0
        last_count = count
