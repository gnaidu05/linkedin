#!/usr/bin/env python3
"""Single entry point for the teammate: `python run.py`.

Prompts for a search term (or takes one as an argument), opens a browser for
you to log into LinkedIn, collects matching posts, and writes the spreadsheet.
"""

from linkedin_posts.cli import main

if __name__ == "__main__":
    raise SystemExit(main())
