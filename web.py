#!/usr/bin/env python3
"""Start the local web page: `python web.py`, then open the printed URL.

Type a search term, click Collect (a browser opens for LinkedIn login), and
download the resulting spreadsheet.
"""

from linkedin_posts.webapp import main

if __name__ == "__main__":
    main()
