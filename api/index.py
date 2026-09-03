"""Public web interface, deployable to Vercel as a serverless function.

Flow: the visitor types a search term and points the page at a LinkedIn
search-results page they saved from their own logged-in browser. The page
extracts the posts *in the browser* and sends only the extracted rows to the
server, which returns a formatted .xlsx.

Why extraction happens in the browser: hosted functions cap request bodies
(4.5 MB on Vercel), and a saved LinkedIn page is far larger than that -- most
of its weight is inline scripts and JSON that we discard anyway. Uploading the
whole file returns 413 FUNCTION_PAYLOAD_TOO_LARGE. Extracting first reduces a
multi-megabyte page to a few KB of JSON, so page size stops mattering.

Why the human supplies the page at all: a public server has no browser for
someone to log into and no LinkedIn session of its own. The only way it could
scrape live results would be to hold LinkedIn credentials, which violates
LinkedIn's Terms of Service. So the human does the LinkedIn part in their own
browser and the server only formats what they hand it.

Stateless by design: nothing is written to disk (serverless filesystems are
ephemeral). The workbook is built in memory and streamed back.
"""

from __future__ import annotations

import os
import sys
from datetime import datetime, timezone
from io import BytesIO

from flask import Flask, jsonify, render_template_string, request, send_file

# Make the repo root importable so the shared package can be reused.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from linkedin_posts.excel import build_workbook  # noqa: E402
from linkedin_posts.parse import parse_posts  # noqa: E402
from linkedin_posts.store import StoredPost  # noqa: E402

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# Bounds on what the JSON endpoint will accept, since it is a public boundary.
MAX_POSTS = 2000
MAX_TEXT_LEN = 20000
MAX_FIELD_LEN = 500

PAGE = r"""<!doctype html>
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
    textarea { min-height: 110px; font-family: ui-monospace, monospace;
               font-size: .85rem; }
    .hint { color: #666; font-size: .85rem; font-weight: 400; margin: .35rem 0 0; }
    .or { text-align: center; color: #888; margin: .75rem 0; font-size: .9rem; }
    button { padding: .7rem 1.3rem; font-size: 1rem; font-weight: 600;
             border: 0; border-radius: 6px; background: #0a66c2; color: #fff;
             cursor: pointer; }
    button:hover { background: #084b8f; }
    button[disabled] { opacity: .6; cursor: progress; }
    .steps { background: #f4f6f8; border: 1px solid #dfe3e8; border-radius: 8px;
             padding: 1rem 1.1rem 1rem 2rem; margin: 0 0 1.5rem; }
    .steps li { margin: .3rem 0; }
    .note { background: #fff8e1; border: 1px solid #ffe082; color: #5f4b00;
            padding: .7rem .9rem; border-radius: 6px; font-size: .9rem;
            margin-bottom: 1.5rem; }
    .msg { padding: .7rem .9rem; border-radius: 6px; margin: 1rem 0 0;
           display: none; }
    .msg.error { display: block; background: #fdecea; border: 1px solid #f5c6cb;
                 color: #8a1c1c; }
    .msg.ok { display: block; background: #e8f5e9; border: 1px solid #a5d6a7;
              color: #1b5e20; }
    .msg.busy { display: block; background: #e3f2fd; border: 1px solid #90caf9;
                color: #0d47a1; }
    code { background: #eceff1; padding: .1rem .3rem; border-radius: 3px;
           font-size: .9em; }
  </style>
</head>
<body>
  <h1>LinkedIn Posts &rarr; Excel</h1>
  <p class="sub">Turn a saved LinkedIn search-results page into a spreadsheet.</p>

  {% if error %}<div class="msg error">{{ error }}</div>{% endif %}

  <ol class="steps">
    <li>In your own browser, log in to LinkedIn and run your search under
        <strong>Posts</strong>.</li>
    <li>Scroll until you have as many posts as you want.</li>
    <li>Save the page: <code>Ctrl/Cmd&nbsp;+&nbsp;S</code>, choosing
        <strong>Webpage, HTML Only</strong>.</li>
    <li>Choose that <code>.html</code> file below.</li>
  </ol>

  <div class="note">
    Your file is read <strong>in your browser</strong> &mdash; only the extracted
    post rows are sent, so page size doesn't matter. This site never connects to
    LinkedIn and stores no credentials or data.
  </div>

  <form id="form" method="post" action="/generate" enctype="multipart/form-data">
    <fieldset>
      <legend>Search term</legend>
      <label for="term">What did you search for?</label>
      <input type="text" id="term" name="term" value="{{ term }}"
             placeholder="e.g. data engineering" required>
      <p class="hint">Recorded in the spreadsheet's "Search Term" column.</p>
    </fieldset>

    <fieldset>
      <legend>Saved search results</legend>
      <label for="html_file">Choose the saved HTML file</label>
      <input type="file" id="html_file" name="html_file" accept=".html,.htm,text/html">
      <div class="or">&mdash; or &mdash;</div>
      <label for="html_text">Paste the page source</label>
      <textarea id="html_text" name="html_text"
                placeholder="Paste the saved page's HTML here instead"></textarea>
    </fieldset>

    <button type="submit" id="go">Collect posts &amp; download Excel</button>
    <div class="msg" id="msg"></div>
  </form>

  <details id="diagwrap" style="display:none;margin-top:1rem">
    <summary style="cursor:pointer;font-weight:600">What the page saw</summary>
    <p class="hint">If the result looks wrong, copy this and send it along &mdash;
       it says what was actually in your file.</p>
    <pre id="diag" style="overflow-x:auto;background:#f4f6f8;border:1px solid #dfe3e8;
         border-radius:6px;padding:.8rem;font-size:.8rem;white-space:pre-wrap"></pre>
  </details>

<script>
/* Extraction runs here, in the browser, so a multi-megabyte saved page never
   crosses the network.

   LinkedIn's class names change and are partly obfuscated, so this does not
   rely on them. It finds posts by their activity URN -- which appears in data
   attributes, permalink hrefs and inline JSON -- then walks up to the
   surrounding card and infers the fields structurally: the author from a
   profile/company link, the date from a <time> or a relative-time string, the
   post body from the largest text block. Known class names are still tried
   first when they happen to be present. */

const KNOWN_CONTAINERS = [
  "div.feed-shared-update-v2",
  "div.update-components-update-v2",
  "li.reusable-search__result-container",
];
const KNOWN_AUTHORS = [
  "span.update-components-actor__title span[aria-hidden='true']",
  "span.update-components-actor__title span.visually-hidden",
  "span.update-components-actor__title",
  "span.update-components-actor__name",
];
const KNOWN_DATES = [
  "span.update-components-actor__sub-description span[aria-hidden='true']",
  "span.update-components-actor__sub-description",
  "time",
];
const KNOWN_TEXTS = [
  "div.update-components-text",
  "div.feed-shared-update-v2__description",
  "span.break-words",
];

const URN_RE = /urn:li:activity:\d+/;
const URN_RE_ALL = /urn:li:activity:\d+/g;
const RELATIVE_TIME =
  /^(now|\d+\s*(s|m|h|d|w|mo|y)|\d+\s*(second|minute|hour|day|week|month|year)s?(\s+ago)?)$/i;

function clean(s) {
  return (s || "").replace(/ /g, " ").replace(/\s+/g, " ").trim()
                  .replace(/^[•·|\s]+/, "").replace(/[•·|\s]+$/, "").trim();
}
function firstText(el, sels) {
  for (const s of sels) {
    let n = null;
    try { n = el.querySelector(s); } catch (_) {}
    if (n) { const t = clean(n.textContent); if (t) return t; }
  }
  return "";
}
function permalink(urn) {
  return "https://www.linkedin.com/feed/update/" + urn + "/";
}

/* Grow the card outward from the node carrying the URN, stopping before an
   ancestor that would take in a second post. That boundary -- one post per
   card -- is what keeps neighbouring posts from bleeding into each other. */
function containsAtLeastTwo(ancestor, bearers) {
  let n = 0;
  for (const b of bearers) {
    if (ancestor === b.el || ancestor.contains(b.el)) { n++; if (n > 1) return true; }
  }
  return false;
}
function climbToCard(el, bearers) {
  let best = el;
  let cur = el;
  for (let i = 0; i < 12 && cur.parentElement; i++) {
    const parent = cur.parentElement;
    if (!parent.tagName || parent.tagName === "BODY" || parent.tagName === "HTML") break;
    if (containsAtLeastTwo(parent, bearers)) break;
    cur = parent;
    best = parent;
  }
  return best;
}

function findAuthor(card) {
  const known = firstText(card, KNOWN_AUTHORS);
  if (known) return known;
  /* A profile or company link's text is the most reliable author signal. */
  let nodes = [];
  try {
    nodes = card.querySelectorAll('a[href*="/in/"], a[href*="/company/"], a[href*="/school/"]');
  } catch (_) {}
  for (const a of nodes) {
    const t = clean(a.textContent);
    if (t && t.length <= 120 && !/^\d+$/.test(t)) return t;
  }
  /* Otherwise the first short line of the card often is the name. */
  const first = clean((card.textContent || "").split("\n")[0]).slice(0, 120);
  return first.length <= 120 ? first : "";
}

function findDate(card) {
  const t0 = card.querySelector("time");
  if (t0) {
    const t = clean(t0.getAttribute("datetime") || t0.textContent);
    if (t) return t;
  }
  const known = firstText(card, KNOWN_DATES);
  if (known && known.length <= 40) return known;
  let leaves = [];
  try { leaves = card.querySelectorAll("span,div,time,p"); } catch (_) {}
  for (const el of leaves) {
    if (el.children && el.children.length) continue;
    const t = clean(el.textContent);
    if (t && t.length <= 24 && RELATIVE_TIME.test(t)) return t;
  }
  return "";
}

function findText(card) {
  const known = firstText(card, KNOWN_TEXTS);
  if (known) return known;
  /* Largest text block that is not the whole card wins. */
  let best = "";
  let nodes = [];
  try { nodes = card.querySelectorAll("div,span,p"); } catch (_) {}
  for (const el of nodes) {
    const t = clean(el.textContent);
    if (t.length > best.length && t.length < (card.textContent || "").length) best = t;
  }
  if (!best) best = clean(card.textContent);
  return best;
}

/* Map post URL -> a DOM node that mentions it. */
function urnNodes(doc) {
  const map = new Map();
  const remember = function (urn, el) {
    const url = permalink(urn);
    if (!map.has(url)) map.set(url, el);
  };
  let all = [];
  try { all = doc.querySelectorAll("*"); } catch (_) {}
  for (const el of all) {
    if (!el.attributes) continue;
    for (const attr of el.attributes) {
      const m = URN_RE.exec(attr.value || "");
      if (m) { remember(m[0], el); break; }
    }
  }
  return map;
}

function extractPosts(html, doc) {
  const out = [];
  const seen = new Set();

  /* 1. Known containers first, when the markup is the familiar shape. */
  for (const sel of KNOWN_CONTAINERS) {
    let els = [];
    try { els = doc.querySelectorAll(sel); } catch (_) {}
    for (const el of els) {
      const m = URN_RE.exec(el.outerHTML || "");
      if (!m) continue;
      const url = permalink(m[0]);
      if (seen.has(url)) continue;
      seen.add(url);
      out.push({ url: url, date_posted: findDate(el), author: findAuthor(el), text: findText(el) });
    }
  }

  /* 2. URN-anywhere: works regardless of class names. */
  const found = urnNodes(doc);
  const bearers = [];
  for (const [u, el] of found) bearers.push({ url: u, el: el });
  for (const b of bearers) {
    if (seen.has(b.url)) continue;
    seen.add(b.url);
    const card = climbToCard(b.el, bearers);
    out.push({
      url: b.url,
      date_posted: findDate(card),
      author: findAuthor(card),
      text: findText(card),
    });
  }

  /* 3. URNs present only in inline JSON: emit link-only rows so nothing is
        silently lost. */
  const inRaw = html.match(URN_RE_ALL) || [];
  for (const urn of inRaw) {
    const url = permalink(urn);
    if (seen.has(url)) continue;
    seen.add(url);
    out.push({ url: url, date_posted: "", author: "", text: "" });
  }
  return out;
}

function diagnose(html, doc) {
  const urns = new Set(html.match(URN_RE_ALL) || []);
  const lines = [];
  lines.push("file size: " + (html.length / 1048576).toFixed(2) + " MB");
  lines.push("activity URNs in file: " + urns.size);
  for (const sel of KNOWN_CONTAINERS) {
    let n = 0;
    try { n = doc.querySelectorAll(sel).length; } catch (_) {}
    lines.push("matches " + sel + ": " + n);
  }
  let permalinks = 0;
  try {
    permalinks = doc.querySelectorAll('a[href*="/feed/update/"], a[href*="/posts/"]').length;
  } catch (_) {}
  lines.push("permalink anchors: " + permalinks);
  const freq = {};
  let classed = [];
  try { classed = doc.querySelectorAll("div[class],li[class],article[class],section[class]"); } catch (_) {}
  for (const el of classed) {
    for (const c of el.classList) freq[c] = (freq[c] || 0) + 1;
  }
  const top = Object.keys(freq).sort(function (a, b) { return freq[b] - freq[a]; }).slice(0, 15);
  lines.push("most common class names:");
  for (const c of top) lines.push("  " + c + " x" + freq[c]);
  return lines.join("\n");
}

const form = document.getElementById("form");
const msg = document.getElementById("msg");
const go = document.getElementById("go");
const diagWrap = document.getElementById("diagwrap");
const diagPre = document.getElementById("diag");
function show(kind, text) { msg.className = "msg " + kind; msg.textContent = text; }

form.addEventListener("submit", async function (e) {
  e.preventDefault();
  diagWrap.style.display = "none";
  const term = document.getElementById("term").value.trim();
  if (!term) { show("error", "Please enter a search term."); return; }

  let html = "";
  const file = document.getElementById("html_file").files[0];
  try {
    if (file) {
      show("busy", "Reading " + file.name + "…");
      html = await file.text();
    }
    if (!html.trim()) html = document.getElementById("html_text").value;
    if (!html.trim()) {
      show("error", "Please choose the saved HTML file, or paste the page source.");
      return;
    }

    show("busy", "Finding posts in the page…");
    const doc = new DOMParser().parseFromString(html, "text/html");
    const posts = extractPosts(html, doc);

    /* Always make the diagnostics available -- it is what turns "not working"
       into something fixable. */
    diagPre.textContent = diagnose(html, doc) + "\nposts extracted: " + posts.length;
    diagWrap.style.display = "block";

    if (!posts.length) {
      show("error", "No posts found in that page. Open “What the page saw” " +
                    "below and send it along so the extractor can be adjusted.");
      return;
    }

    go.disabled = true;
    show("busy", "Found " + posts.length + " post(s). Building spreadsheet…");
    const resp = await fetch("/generate-json", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ term: term, posts: posts }),
    });
    if (!resp.ok) {
      let detail = "";
      try { detail = (await resp.json()).error || ""; } catch (_) {}
      show("error", detail || ("Server returned " + resp.status + "."));
      return;
    }
    const blob = await resp.blob();
    const name = resp.headers.get("X-Filename") || "linkedin_posts.xlsx";
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 10000);
    const withText = posts.filter(function (p) { return p.text; }).length;
    show("ok", "Done — " + posts.length + " post(s) in " + name +
               (withText < posts.length
                 ? ". " + (posts.length - withText) + " had only a link; open " +
                   "“What the page saw” if that looks wrong."
                 : "."));
  } catch (err) {
    show("error", "Could not process that file: " + err.message);
  } finally {
    go.disabled = false;
  }
});
</script>
</body>
</html>
"""

app = Flask(__name__)
# The JSON path keeps bodies tiny; this only bounds the no-JS upload fallback.
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024


def _render(error: str | None = None, term: str = "", status: int = 200):
    return render_template_string(PAGE, error=error, term=term), status


def _spreadsheet_response(stored: list[StoredPost], term: str):
    buffer = BytesIO()
    build_workbook(stored).save(buffer)
    buffer.seek(0)
    filename = f"{_safe_filename(term)}.xlsx"
    resp = send_file(
        buffer, mimetype=XLSX_MIME, as_attachment=True, download_name=filename
    )
    # Same-origin fetch reads this to name the downloaded file.
    resp.headers["X-Filename"] = filename
    return resp


def _to_stored(records: list[dict], term: str) -> list[StoredPost]:
    """Convert extracted records to rows, de-duplicating by post URL."""
    captured_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    seen: set[str] = set()
    stored: list[StoredPost] = []
    for rec in records:
        url = str(rec.get("url") or "").strip()[:MAX_FIELD_LEN]
        if not url or url in seen:
            continue
        seen.add(url)
        stored.append(
            StoredPost(
                url=url,
                date_posted=str(rec.get("date_posted") or "")[:MAX_FIELD_LEN],
                author=str(rec.get("author") or "")[:MAX_FIELD_LEN],
                text=str(rec.get("text") or "")[:MAX_TEXT_LEN],
                search_term=term,
                captured_at=captured_at,
            )
        )
    return stored


@app.get("/")
def index():
    return _render()


@app.post("/generate-json")
def generate_json():
    """Build a spreadsheet from rows the page already extracted."""
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify(error="Expected a JSON object."), 400

    term = str(payload.get("term") or "").strip()[:MAX_FIELD_LEN]
    if not term:
        return jsonify(error="Please enter a search term."), 400

    records = payload.get("posts")
    if not isinstance(records, list) or not records:
        return jsonify(error="No posts were sent."), 400
    if len(records) > MAX_POSTS:
        return jsonify(error=f"Too many posts (limit {MAX_POSTS})."), 400
    if not all(isinstance(r, dict) for r in records):
        return jsonify(error="Each post must be an object."), 400

    stored = _to_stored(records, term)
    if not stored:
        return jsonify(error="None of the posts had a usable link."), 400
    return _spreadsheet_response(stored, term)


@app.post("/generate")
def generate():
    """No-JavaScript fallback: parse uploaded HTML server-side.

    Subject to the host's request-size limit, so the in-browser path above is
    the primary route for real saved pages.
    """
    term = (request.form.get("term") or "").strip()
    if not term:
        return _render(error="Please enter a search term.", status=400)

    html = ""
    upload = request.files.get("html_file")
    if upload is not None and upload.filename:
        html = upload.read().decode("utf-8", errors="replace")
    if not html.strip():
        html = request.form.get("html_text") or ""
    if not html.strip():
        return _render(
            error="Please choose the saved HTML file, or paste the page source.",
            term=term,
            status=400,
        )

    posts = parse_posts(html)
    if not posts:
        return _render(
            error=(
                "No posts found in that HTML. Make sure you saved the LinkedIn "
                "search results page (the Posts tab) while logged in."
            ),
            term=term,
            status=400,
        )

    stored = _to_stored(
        [
            {
                "url": p.url,
                "date_posted": p.date_posted,
                "author": p.author,
                "text": p.text,
            }
            for p in posts
        ],
        term,
    )
    return _spreadsheet_response(stored, term)


@app.errorhandler(413)
def too_large(_err):
    return _render(
        error=(
            "That file is too large to upload. Enable JavaScript and it will be "
            "read in your browser instead, with no size limit."
        ),
        status=413,
    )


def _safe_filename(term: str) -> str:
    keep = [c if c.isalnum() or c in "-_" else "_" for c in term.strip()]
    stem = "".join(keep).strip("_") or "linkedin_posts"
    return f"linkedin_posts_{stem}"[:80]


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5001, debug=False)
