"""SQLite datastore: the source of truth for collected posts.

The .xlsx is a generated view of this table, never the other way around.
De-duplication is by post URL (the primary key), so re-running the same search
term appends only genuinely new posts.
"""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from .parse import Post

SCHEMA = """
CREATE TABLE IF NOT EXISTS posts (
    url         TEXT PRIMARY KEY,
    date_posted TEXT,
    author      TEXT,
    text        TEXT,
    search_term TEXT,
    captured_at TEXT NOT NULL
);
"""


@dataclass
class StoreResult:
    """Outcome of storing one batch of posts."""

    new_count: int
    duplicate_count: int
    total_count: int


@dataclass
class StoredPost:
    """A row as it lives in the datastore (includes capture metadata)."""

    url: str
    date_posted: str
    author: str
    text: str
    search_term: str
    captured_at: str


class Store:
    def __init__(self, db_path: str | Path):
        self.db_path = str(db_path)
        self._conn = sqlite3.connect(self.db_path)
        self._conn.row_factory = sqlite3.Row
        self._conn.execute(SCHEMA)
        self._conn.commit()

    def close(self) -> None:
        self._conn.close()

    def __enter__(self) -> "Store":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    def add_posts(self, posts: list[Post], search_term: str) -> StoreResult:
        """Insert posts, skipping any URL already present.

        Returns counts of newly inserted vs. skipped duplicates, plus the new
        grand total in the store.
        """
        captured_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
        new_count = 0
        duplicate_count = 0

        for post in posts:
            cur = self._conn.execute(
                """
                INSERT OR IGNORE INTO posts
                    (url, date_posted, author, text, search_term, captured_at)
                VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    post.url,
                    post.date_posted,
                    post.author,
                    post.text,
                    search_term,
                    captured_at,
                ),
            )
            if cur.rowcount == 1:
                new_count += 1
            else:
                duplicate_count += 1

        self._conn.commit()
        return StoreResult(
            new_count=new_count,
            duplicate_count=duplicate_count,
            total_count=self.count(),
        )

    def count(self) -> int:
        row = self._conn.execute("SELECT COUNT(*) AS n FROM posts").fetchone()
        return int(row["n"])

    def all_posts(self) -> list[StoredPost]:
        """Return every stored post, newest capture first."""
        rows = self._conn.execute(
            """
            SELECT url, date_posted, author, text, search_term, captured_at
            FROM posts
            ORDER BY captured_at DESC, url ASC
            """
        ).fetchall()
        return [
            StoredPost(
                url=r["url"],
                date_posted=r["date_posted"] or "",
                author=r["author"] or "",
                text=r["text"] or "",
                search_term=r["search_term"] or "",
                captured_at=r["captured_at"] or "",
            )
            for r in rows
        ]
