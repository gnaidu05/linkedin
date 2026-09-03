"""Tests for the web front end, wired to a fixture instead of a browser."""

from __future__ import annotations

from pathlib import Path

from linkedin_posts.config import Config
from linkedin_posts.webapp import create_app

FIXTURE = Path(__file__).parent / "fixtures" / "sample_search.html"


def _app(tmp_path):
    config = Config(
        db_path=str(tmp_path / "web.db"),
        xlsx_path=str(tmp_path / "web.xlsx"),
    )
    html = FIXTURE.read_text()
    return create_app(config=config, html_provider=lambda term, c: html), config


def test_index_loads_empty(tmp_path):
    app, _ = _app(tmp_path)
    client = app.test_client()
    resp = client.get("/")
    assert resp.status_code == 200
    assert b"LinkedIn Posts" in resp.data
    assert b"No posts collected yet." in resp.data


def test_empty_term_is_rejected(tmp_path):
    app, _ = _app(tmp_path)
    client = app.test_client()
    resp = client.post("/run", data={"term": "   "})
    assert resp.status_code == 400
    assert b"Please enter a search term." in resp.data


def test_run_collects_and_offers_download(tmp_path):
    app, config = _app(tmp_path)
    client = app.test_client()

    resp = client.post("/run", data={"term": "data engineering"})
    assert resp.status_code == 200
    # Summary reflects 3 unique posts from the fixture (one dup URN skipped).
    assert b"3</strong> total" in resp.data
    assert b"Ada Lovelace" in resp.data
    assert Path(config.xlsx_path).exists()

    # The generated spreadsheet downloads.
    dl = client.get("/download")
    assert dl.status_code == 200
    assert dl.headers["Content-Disposition"].startswith("attachment")

    # Re-running the same term adds nothing new (safe re-run through the UI).
    resp2 = client.post("/run", data={"term": "data engineering"})
    assert b"0</strong> new added" in resp2.data


def test_post_text_is_escaped(tmp_path):
    """Untrusted post text must be HTML-escaped in the rendered page.

    The source encodes ``<b>bold</b> & co`` as HTML entities, so it survives
    parsing as literal text (rather than being stripped as a real tag). The
    rendered page must escape it again rather than emit a live <b> element.
    """
    config = Config(
        db_path=str(tmp_path / "x.db"), xlsx_path=str(tmp_path / "x.xlsx")
    )
    payload = (
        '<div class="feed-shared-update-v2" '
        'data-urn="urn:li:activity:9000000000000000001">'
        '<span class="update-components-actor__title">'
        '<span aria-hidden="true">Attacker</span></span>'
        '<div class="update-components-text">'
        "&lt;b&gt;bold&lt;/b&gt; &amp; co</div></div>"
    )
    app = create_app(config=config, html_provider=lambda t, c: payload)
    client = app.test_client()
    resp = client.post("/run", data={"term": "xss"})
    assert resp.status_code == 200
    # Rendered as escaped text, not as a live element.
    assert b"<b>bold</b>" not in resp.data
    assert b"&lt;b&gt;bold&lt;/b&gt;" in resp.data


if __name__ == "__main__":
    import subprocess
    import sys

    raise SystemExit(
        subprocess.call([sys.executable, "-m", "pytest", "-q", __file__])
    )
