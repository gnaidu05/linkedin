"""A minimal local web page for running a collection.

A text box to type the search term, a button to run it, a table of collected
posts, and a link to download the spreadsheet. It reuses the same pipeline as
the command line -- collect -> parse -> store -> excel -- so it is a thin UI
over existing code, not a second implementation.

Clicking "Collect" still opens the browser window for manual LinkedIn login;
the web page is only the local front end. Run with:  python web.py
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable

from flask import Flask, abort, render_template_string, request, send_file

from .config import Config, load_config
from .excel import write_xlsx
from .parse import parse_posts
from .store import Store, StoredPost

HtmlProvider = Callable[[str, Config], str]

PAGE = """<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>LinkedIn Posts → Excel</title>
  <style>
    :root { color-scheme: light dark; }
    body { font: 15px/1.5 system-ui, sans-serif; margin: 0; padding: 2rem;
           max-width: 1000px; margin-inline: auto; }
    h1 { font-size: 1.4rem; margin: 0 0 .25rem; }
    .sub { color: #666; margin: 0 0 1.5rem; }
    form { display: flex; gap: .5rem; margin-bottom: 1rem; }
    input[type=text] { flex: 1; padding: .6rem .7rem; font-size: 1rem;
                       border: 1px solid #999; border-radius: 6px; }
    button { padding: .6rem 1.1rem; font-size: 1rem; border: 0;
             border-radius: 6px; background: #0a66c2; color: #fff;
             cursor: pointer; }
    button:hover { background: #084b8f; }
    .note { background: #fff8e1; border: 1px solid #ffe082; color: #5f4b00;
            padding: .6rem .8rem; border-radius: 6px; font-size: .9rem;
            margin-bottom: 1.5rem; }
    .error { background: #fdecea; border: 1px solid #f5c6cb; color: #8a1c1c;
             padding: .6rem .8rem; border-radius: 6px; margin-bottom: 1rem; }
    .summary { background: #e8f5e9; border: 1px solid #a5d6a7; color: #1b5e20;
               padding: .7rem .9rem; border-radius: 6px; margin-bottom: 1rem; }
    .toolbar { margin-bottom: 1rem; }
    a.download { display: inline-block; padding: .5rem .9rem; border-radius: 6px;
                 background: #1b5e20; color: #fff; text-decoration: none; }
    table { border-collapse: collapse; width: 100%; font-size: .9rem; }
    th, td { border: 1px solid #ccc; padding: .5rem .6rem; text-align: left;
             vertical-align: top; }
    th { background: #f0f0f0; }
    td.text { max-width: 480px; }
    .count { color: #666; margin: .5rem 0 1rem; }
  </style>
</head>
<body>
  <h1>LinkedIn Posts → Excel</h1>
  <p class="sub">Type a search term and collect matching posts into a spreadsheet.</p>

  <div class="note">
    When you click <strong>Collect</strong>, a browser window opens. Log in to
    LinkedIn there if you aren't already; collection starts automatically.
    Automating LinkedIn is against its Terms of Service — use sparingly.
  </div>

  {% if error %}<div class="error">{{ error }}</div>{% endif %}

  <form method="post" action="/run">
    <input type="text" name="term" value="{{ term }}"
           placeholder="e.g. data engineering" autofocus>
    <button type="submit">Collect</button>
  </form>

  {% if summary %}
  <div class="summary">
    Run complete for “{{ term }}”: found {{ summary.found }} on the page —
    <strong>{{ summary.new }}</strong> new added,
    <strong>{{ summary.dup }}</strong> duplicate(s) skipped,
    <strong>{{ summary.total }}</strong> total in the datastore.
  </div>
  {% endif %}

  <div class="toolbar">
    {% if xlsx_ready %}
      <a class="download" href="/download">⬇ Download Excel</a>
    {% endif %}
  </div>

  <p class="count">{{ posts|length }} post(s) stored.</p>

  <table>
    <thead>
      <tr><th>Post Link</th><th>Date Posted</th><th>Posted By</th>
          <th>Post Text</th><th>Search Term</th><th>Captured At</th></tr>
    </thead>
    <tbody>
      {% for p in posts %}
      <tr>
        <td><a href="{{ p.url }}" target="_blank" rel="noopener">open</a></td>
        <td>{{ p.date_posted }}</td>
        <td>{{ p.author }}</td>
        <td class="text">{{ p.text }}</td>
        <td>{{ p.search_term }}</td>
        <td>{{ p.captured_at }}</td>
      </tr>
      {% else %}
      <tr><td colspan="6">No posts collected yet.</td></tr>
      {% endfor %}
    </tbody>
  </table>
</body>
</html>
"""


def _stored_posts(cfg: Config) -> list[StoredPost]:
    with Store(cfg.db_path) as store:
        return store.all_posts()


def create_app(
    config: Config | None = None, html_provider: HtmlProvider | None = None
) -> Flask:
    """Build the Flask app. ``html_provider`` is the LinkedIn seam (injectable
    for tests); it defaults to the live browser collector."""
    app = Flask(__name__)
    cfg = config or load_config()

    if html_provider is None:

        def html_provider(term: str, c: Config) -> str:  # noqa: E306
            from .collect import collect_html

            return collect_html(term, c)

    @app.get("/")
    def index():
        return render_template_string(
            PAGE,
            posts=_stored_posts(cfg),
            summary=None,
            term="",
            error=None,
            xlsx_ready=Path(cfg.xlsx_path).exists(),
        )

    @app.post("/run")
    def run_collection():
        term = (request.form.get("term") or "").strip()
        if not term:
            return (
                render_template_string(
                    PAGE,
                    posts=_stored_posts(cfg),
                    summary=None,
                    term="",
                    error="Please enter a search term.",
                    xlsx_ready=Path(cfg.xlsx_path).exists(),
                ),
                400,
            )

        html = html_provider(term, cfg)
        parsed = parse_posts(html)
        with Store(cfg.db_path) as store:
            result = store.add_posts(parsed, term)
            posts = store.all_posts()
            write_xlsx(posts, cfg.xlsx_path)

        summary = {
            "found": len(parsed),
            "new": result.new_count,
            "dup": result.duplicate_count,
            "total": result.total_count,
        }
        return render_template_string(
            PAGE,
            posts=posts,
            summary=summary,
            term=term,
            error=None,
            xlsx_ready=True,
        )

    @app.get("/download")
    def download():
        path = Path(cfg.xlsx_path)
        if not path.exists():
            abort(404)
        return send_file(
            path.resolve(), as_attachment=True, download_name=path.name
        )

    return app


def main() -> None:
    app = create_app()
    print("Open http://127.0.0.1:5000 in your browser.")
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
