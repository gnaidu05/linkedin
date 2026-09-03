"""Runtime configuration, loaded from config.json with sane defaults."""

from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path

DEFAULT_CONFIG_PATH = Path("config.json")


@dataclass
class Config:
    # How many posts to try to collect in one run.
    max_results: int = 50
    # Where the source-of-truth datastore lives.
    db_path: str = "linkedin_posts.db"
    # Where the generated spreadsheet is written.
    xlsx_path: str = "linkedin_posts.xlsx"
    # Browser profile directory. Reusing one lets the teammate stay logged in
    # across runs (their own session, like a normal browser profile). It holds
    # no credentials we set -- LinkedIn's own session cookies, created when the
    # human logs in by hand.
    user_data_dir: str = ".browser-profile"
    # Seconds to wait for the human to finish logging in before giving up.
    login_timeout_seconds: int = 300
    # Seconds to pause between scrolls while lazy-loaded posts stream in.
    scroll_pause_seconds: float = 2.0


def load_config(path: Path | str = DEFAULT_CONFIG_PATH) -> Config:
    """Load config from JSON, falling back to defaults for any missing key."""
    path = Path(path)
    if not path.exists():
        return Config()
    data = json.loads(path.read_text())
    defaults = Config()
    return Config(
        max_results=int(data.get("max_results", defaults.max_results)),
        db_path=str(data.get("db_path", defaults.db_path)),
        xlsx_path=str(data.get("xlsx_path", defaults.xlsx_path)),
        user_data_dir=str(data.get("user_data_dir", defaults.user_data_dir)),
        login_timeout_seconds=int(
            data.get("login_timeout_seconds", defaults.login_timeout_seconds)
        ),
        scroll_pause_seconds=float(
            data.get("scroll_pause_seconds", defaults.scroll_pause_seconds)
        ),
    )
