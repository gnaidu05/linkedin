"""Public web interface, deployable to Vercel as a serverless function.

Flow: the visitor types a search term, uploads (or pastes) the HTML of a
LinkedIn search-results page they saved from their own logged-in browser, and
gets back a formatted .xlsx.

Why it works this way: a public server has no browser for a human to log into
and no LinkedIn session of its own. The only way it could scrape live results
would be to hold LinkedIn credentials, which violates LinkedIn's Terms of
Service. So the human does the LinkedIn part in their own browser and the
server only parses what they hand it -- no credentials, no automation against
LinkedIn.

Stateless by design: nothing is written to disk (serverless filesystems are
ephemeral). The workbook is built in memory and streamed back. De-duplication
by post URL happens within the submitted HTML.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime, timezone
from io import BytesIO

from flask import Flask, render_template_string, request, send_file

# Make the repo root importable so the shared package can be reused.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from linkedin_posts.excel import build_workbook  # noqa: E402
from linkedin_posts.parse import parse_posts  # noqa: E402
from linkedin_posts.store import StoredPost  # noqa: E402

XLSX_MIME = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

PAGE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>LinkedIn Posts &rarr; Excel</title>
  <style>
    :root { color-scheme: light dark; }
    * { box-sizing: border-box; }
    body { font: 15px/1.6 system-ui, -apple-system, sans-serif; margin: 0;
           padding: 2rem 1.25rem; max-width: 760px; margin-inline: auto; }
    h1 { font-size: 1.5rem; margin: 0 0 .25rem; }
    .sub { color: #666; margin: 0 0 1.75rem; }
    fieldset { border: 1px solid #ccc; border-radius: 8px; padding: 1rem 1.1rem;
               margin: 0 0 1rem; }
    legend { font-weight: 600; padding: 0 .4rem; }
    label { display: block; font-weight: 600; margin-bottom: .35rem; }
    input[type=text], textarea, input[type=file] { width: 100%;
      padding: .6rem .7rem; font: inherit; border: 1px solid #999;
      border-radius: 6px; background: canvas; color: canvastext; }
    textarea { min-height: 120px; font-family: ui-monospace, monospace;
               font-size: .85rem; }
    .hint { color: #666; font-size: .85rem; font-weight: 400;
            margin: .35rem 0 0; }
    .or { text-align: center; color: #888; margin: .75rem 0; font-size: .9rem; }
    button { padding: .7rem 1.3rem; font-size: 1rem; font-weight: 600;
             border: 0; border-radius: 6px; background: #0a66c2; color: #fff;
             cursor: pointer; }
    button:hover { background: #084b8f; }
    .steps { background: #f4f6f8; border: 1px solid #dfe3e8; border-radius: 8px;
             padding: 1rem 1.1rem 1rem 2rem; margin: 0 0 1.5rem; }
    .steps li { margin: .3rem 0; }
    .note { background: #fff8e1; border: 1px solid #ffe082; color: #5f4b00;
            padding: .7rem .9rem; border-radius: 6px; font-size: .9rem;
            margin-bottom: 1.5rem; }
    .error { background: #fdecea; border: 1px solid #f5c6cb; color: #8a1c1c;
             padding: .7rem .9rem; border-radius: 6px; margin-bottom: 1rem; }
    code { background: #eceff1; padding: .1rem .3rem; border-radius: 3px;
           font-size: .9em; }
  </style>
</head>
<body>
  <h1>LinkedIn Posts &rarr; Excel</h1>
  <p class="sub">Turn a saved LinkedIn search-results page into a spreadsheet.</p>

  {% if error %}<div class="error">{{ error }}</div>{% endif %}

  <ol class="steps">
    <li>In your own browser, log in to LinkedIn and run your search under
        <strong>Posts</strong>.</li>
    <li>Scroll until you have as many posts as you want.</li>
    <li>Save the page: <code>Ctrl/Cmd&nbsp;+&nbsp;S</code>, choosing
        <strong>Webpage, HTML Only</strong> (not "Complete").</li>
    <li>Upload that <code>.html</code> file below.</li>
  </ol>

  <div class="note">
    This site never connects to LinkedIn and stores no credentials or data &mdash;
    it only reads the file you upload and returns a spreadsheet.
  </div>

  <form method="post" action="/generate" enctype="multipart/form-data">
    <fieldset>
      <legend>Search term</legend>
      <label for="term">What did you search for?</label>
      <input type="text" id="term" name="term" value="{{ term }}"
             placeholder="e.g. data engineering" required>
      <p class="hint">Recorded in the spreadsheet's "Search Term" column.</p>
    </fieldset>

    <fieldset>
      <legend>Saved search results</legend>
      <label for="html_file">Upload the saved HTML file</label>
      <input type="file" id="html_file" name="html_file" accept=".html,.htm,text/html">
      <div class="or">&mdash; or &mdash;</div>
      <label for="html_text">Paste the page source</label>
      <textarea id="html_text" name="html_text"
                placeholder="Paste the saved page's HTML here instead"></textarea>
    </fieldset>

    <button type="submit">Collect posts &amp; download Excel</button>
  </form>
</body>
</html>
"""

app = Flask(__name__)
# Keep uploads bounded; serverless request bodies are limited anyway.
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB


def _render(error: str | None = None, term: str = "", status: int = 200):
    return render_template_string(PAGE, error=error, term=term), status


@app.get("/")
def index():
    return _render()


@app.post("/generate")
def generate():
    term = (request.form.get("term") or "").strip()
    if not term:
        return _render(error="Please enter a search term.", status=400)

    html = ""
    upload = request.files.get("html_file")
    if upload is not None and upload.filename:
        raw = upload.read()
        html = raw.decode("utf-8", errors="replace")
    if not html.strip():
        html = request.form.get("html_text") or ""
    if not html.strip():
        return _render(
            error="Please upload the saved HTML file, or paste the page source.",
            term=term,
            status=400,
        )

    posts = parse_posts(html)
    if not posts:
        return _render(
            error=(
                "No posts found in that HTML. Make sure you saved the LinkedIn "
                "search results page (the Posts tab) while logged in, using "
                '"Webpage, HTML Only".'
            ),
            term=term,
            status=400,
        )

    # De-duplicate by post URL, keeping first occurrence order.
    captured_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    seen: set[str] = set()
    stored: list[StoredPost] = []
    for post in posts:
        if post.url in seen:
            continue
        seen.add(post.url)
        stored.append(
            StoredPost(
                url=post.url,
                date_posted=post.date_posted,
                author=post.author,
                text=post.text,
                search_term=term,
                captured_at=captured_at,
            )
        )

    buffer = BytesIO()
    build_workbook(stored).save(buffer)
    buffer.seek(0)
    return send_file(
        buffer,
        mimetype=XLSX_MIME,
        as_attachment=True,
        download_name=f"{_safe_filename(term)}.xlsx",
    )


def _safe_filename(term: str) -> str:
    """Turn a search term into a conservative filename stem."""
    keep = [c if c.isalnum() or c in "-_" else "_" for c in term.strip()]
    stem = "".join(keep).strip("_") or "linkedin_posts"
    return f"linkedin_posts_{stem}"[:80]


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5001, debug=False)
