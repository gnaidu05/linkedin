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

If ToS-compliance matters more than live search, the safer alternative is to
save the search-results page as HTML from your own browser and parse that
offline — ask and this can be adapted to that.

---

## Setup

Requires Python 3.11+.

```bash
pip install -r requirements.txt
playwright install chromium
```

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
run.py                     One-command entry point.
config.json                Settings.
linkedin_posts/
  cli.py                   Orchestrates a run: collect -> parse -> store -> excel.
  collect.py               The only LinkedIn-facing code: drives the browser.
  parse.py                 Pure HTML -> post records (no network; unit-tested).
  store.py                 SQLite datastore; de-dup and incremental append.
  excel.py                 Builds the formatted .xlsx from the datastore.
  config.py                Loads config.json.
tests/
  test_pipeline.py         Parse, store, excel, and full pipeline (on a fixture).
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
