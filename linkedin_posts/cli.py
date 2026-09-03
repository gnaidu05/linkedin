"""Command-line entry point: search term in, Excel file out."""

from __future__ import annotations

import argparse
import sys
from typing import Callable

from .config import Config, load_config
from .excel import write_xlsx
from .parse import parse_posts
from .store import Store

# A function that, given a search term and config, returns page HTML. The real
# one drives a browser; tests inject a fixture. This is the pipeline's only
# seam onto LinkedIn.
HtmlProvider = Callable[[str, Config], str]


def run(search_term: str, config: Config, html_provider: HtmlProvider) -> int:
    """Run the full pipeline once. Returns a process exit code."""
    search_term = search_term.strip()
    if not search_term:
        print("Error: search term must not be empty.", file=sys.stderr)
        return 2

    print(f'Collecting posts for: "{search_term}"')
    html = html_provider(search_term, config)
    posts = parse_posts(html)
    print(f"Found {len(posts)} post(s) on the page.")

    with Store(config.db_path) as store:
        result = store.add_posts(posts, search_term)
        write_xlsx(store.all_posts(), config.xlsx_path)

    # End-of-run summary.
    print(
        "\nRun complete:\n"
        f"  {result.new_count} new post(s) added\n"
        f"  {result.duplicate_count} duplicate(s) skipped\n"
        f"  {result.total_count} post(s) total in the datastore "
        f"({config.db_path})\n"
        f"  Spreadsheet written to {config.xlsx_path}"
    )
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="linkedin-posts",
        description="Collect LinkedIn posts matching a search term into Excel.",
    )
    parser.add_argument(
        "search_term",
        nargs="?",
        help="Search term. If omitted, you will be prompted for it.",
    )
    parser.add_argument(
        "--config",
        default="config.json",
        help="Path to config JSON (default: config.json).",
    )
    args = parser.parse_args(argv)

    config = load_config(args.config)

    search_term = args.search_term
    if not search_term:
        try:
            search_term = input("Enter a search term: ")
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.", file=sys.stderr)
            return 130

    # Imported here so `--help` and prompting work without Playwright installed.
    from .collect import collect_html

    return run(search_term, config, collect_html)


if __name__ == "__main__":
    raise SystemExit(main())
