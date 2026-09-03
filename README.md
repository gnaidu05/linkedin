# LinkedIn Posts → Excel

Collect LinkedIn posts matching a search term into an Excel spreadsheet. Run
one command, type a search term, and get back an `.xlsx` file where each row is
one post.

Each row has: **post link** (clickable), **date posted**, **posted by**
(author), and the **post text** — plus the search term used and a timestamp of
when the post was captured. The header row is bold and frozen, columns are
auto-sized, and the post-link column contains working hyperlinks.

---

## ⚠️ Please read first: how this accesses LinkedIn, and the risk

There is no official LinkedIn API that lets you keyword-search public posts, so
this tool opens a real browser window and you **log in to LinkedIn yourself, by
hand.** The tool then runs your search and reads the results off the page.

- **No passwords are stored or typed by this tool.** You log in in the browser
  window, exactly as you would normally. LinkedIn's own session cookies are
  kept in a local `.browser-profile/` folder (like a normal browser profile) so
  you don't have to log in every single run. That folder is git-ignored.
- **Automating LinkedIn's interface is against LinkedIn's Terms of Service**,
  and doing it can get an account rate-limited or restricted. You chose this
  method knowingly. Use a low `max_results`, don't run it in a tight loop, and
  understand the account risk is yours.

If ToS-compliance matters more than live automation, use the
[public web page](#public-web-page-vercel) instead: you save the search-results
page from your own browser and it only parses that file, never touching
LinkedIn itself.

---

## Setup

Requires Python 3.11+.

```bash
pip install -r requirements-local.txt
playwright install chromium
```

(`requirements-local.txt` adds Playwright on top of `requirements.txt`.
`requirements.txt` alone is the smaller set the hosted web page needs, and is
what Vercel installs.)

(`playwright install chromium` downloads the browser Playwright drives. You
only need to do it once.)

## Run

```bash
python run.py
```

It will:

1. Prompt you for a search term (you can also pass it directly:
   `python run.py "data engineering"`).
2. Open a browser window. **Log in to LinkedIn there if you aren't already.**
3. Run the search, scroll to load results, and read the posts.
4. Append any new posts to the local datastore and rebuild the spreadsheet.

When it finishes it prints a summary like:

```
Run complete:
  7 new post(s) added
  3 duplicate(s) skipped
  42 post(s) total in the datastore (linkedin_posts.db)
  Spreadsheet written to linkedin_posts.xlsx
```

Open `linkedin_posts.xlsx` to see the results.

## Or use the web page

If you'd rather work in a browser than a terminal, start the local web page:

```bash
python web.py
```

Then open **http://127.0.0.1:5000**. Type a search term, click **Collect**
(a browser window opens for the manual LinkedIn login, exactly as above), and
when it finishes the page shows a summary and the collected posts, with a
**Download Excel** button. It runs against the same datastore as `run.py`, so
the two entry points share one source of truth.

The web page runs only on your own machine (`127.0.0.1`); it is a convenience
front end, not a hosted service, and it does not change the LinkedIn login or
Terms-of-Service realities described above.

## Public web page (Vercel)

There is also a **public** version of the page, deployable to Vercel, at
`api/index.py`. It works differently from the local tool, for a reason worth
understanding:

A public server has no browser for you to log into and no LinkedIn session of
its own. The only way it could scrape live results would be to store LinkedIn
credentials on the server, which violates LinkedIn's Terms of Service. So the
public page splits the work: **you** do the LinkedIn part in your own browser,
and the server only processes what you hand it.

How your teammate uses it:

1. Log in to LinkedIn in their own browser and run the search under **Posts**.
2. Scroll until enough posts are loaded.
3. Save the page with `Ctrl/Cmd + S`, choosing **Webpage, HTML Only**
   (not "Complete" -- HTML Only is much smaller).
4. On the public page: type the search term, upload that `.html` file, click
   **Collect posts & download Excel**.

The page reads your file **in your browser**, extracts the posts, and sends
only those rows to the server, which returns the same formatted `.xlsx`
(frozen bold header, auto-sized columns, clickable links) with duplicates
removed by post URL.

That split matters: hosted functions cap request bodies (4.5 MB on Vercel) and
a real saved LinkedIn page is much bigger than that -- mostly inline scripts
and JSON that get discarded anyway. Uploading the whole file returns
`413 FUNCTION_PAYLOAD_TOO_LARGE`. Extracting first turns a multi-megabyte page
into a few KB of JSON, so page size stops mattering. (A no-JavaScript fallback
still posts the raw HTML to the server, but it is subject to that size limit.)

The site **stores nothing** -- no credentials, no database, no uploaded files;
the workbook is built in memory. That also means it keeps no incremental
history: each run produces a standalone spreadsheet. The local tool is the one
with a growing datastore.

### Deploying

`vercel.json` routes all traffic to the single Python function and Vercel
installs `requirements.txt`. Connect the repo in Vercel and it deploys on push;
no environment variables or secrets are needed.

If the deployment URL asks you to log in to Vercel, that is Vercel's
Deployment Protection, not the app: turn it off under
**Settings -> Deployment Protection -> Vercel Authentication**.

## How re-running works

- The **source of truth** is the local SQLite datastore (`linkedin_posts.db`).
  The `.xlsx` is just a generated view of it, rebuilt from scratch every run.
- Posts are de-duplicated **by post URL**. Re-running the same search term
  **appends** genuinely new posts to what you already have — it never
  overwrites or loses earlier results. Running twice with no new posts is safe
  and simply adds nothing.

## Configuration

Edit `config.json`:

| Key | Meaning | Default |
|-----|---------|---------|
| `max_results` | How many posts to try to collect per run | `50` |
| `db_path` | SQLite datastore (source of truth) | `linkedin_posts.db` |
| `xlsx_path` | Generated spreadsheet | `linkedin_posts.xlsx` |
| `user_data_dir` | Browser profile folder (keeps you logged in) | `.browser-profile` |
| `login_timeout_seconds` | How long to wait for you to log in | `300` |
| `scroll_pause_seconds` | Pause between scrolls while posts load | `2.0` |

## Project layout

```
run.py                     One-command terminal entry point.
web.py                     Local web-page entry point.
api/index.py               Public web page (Vercel serverless function).
vercel.json                Vercel routing/build config.
config.json                Settings.
linkedin_posts/
  cli.py                   Orchestrates a run: collect -> parse -> store -> excel.
  webapp.py                Flask web page over the same pipeline.
  collect.py               The only LinkedIn-facing code: drives the browser.
  parse.py                 Pure HTML -> post records (no network; unit-tested).
  store.py                 SQLite datastore; de-dup and incremental append.
  excel.py                 Builds the formatted .xlsx from the datastore.
  config.py                Loads config.json.
tests/
  test_pipeline.py         Parse, store, excel, and full pipeline (on a fixture).
  test_webapp.py           Local web routes (on a fixture; no browser).
  test_public_api.py       Public upload interface (on a fixture).
  fixtures/sample_search.html
```

## Tests

The parsing, storage, and Excel export are tested end-to-end against a saved
HTML fixture (no LinkedIn or browser needed):

```bash
python -m pytest -q
```

The live browser/login step cannot be exercised without a real logged-in
LinkedIn session, so it is deliberately isolated in `collect.py`; everything
downstream of it is covered by the tests.
