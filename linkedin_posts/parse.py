"""Pure HTML -> post records.

This module contains no network or browser code. It takes the HTML of a
LinkedIn search-results page (as captured by the browser layer) and extracts
one record per post. Keeping it pure means it can be tested against a saved
HTML fixture without touching LinkedIn.

LinkedIn's markup changes over time and varies by surface, so every field
uses a small list of selector fallbacks and degrades to an empty string
rather than raising.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from bs4 import BeautifulSoup
from bs4.element import Tag

# A post card in the feed / search results. LinkedIn has used a few container
# class names over the years; we accept any of them.
POST_CONTAINER_SELECTORS = [
    "div.feed-shared-update-v2",
    "div.update-components-update-v2",
    "li.reusable-search__result-container",
]

AUTHOR_SELECTORS = [
    "span.update-components-actor__title span[aria-hidden='true']",
    "span.update-components-actor__title span.visually-hidden",
    "span.update-components-actor__title",
    "span.update-components-actor__name",
]

DATE_SELECTORS = [
    "span.update-components-actor__sub-description span[aria-hidden='true']",
    "span.update-components-actor__sub-description",
    "time",
]

TEXT_SELECTORS = [
    "div.update-components-text",
    "div.feed-shared-update-v2__description",
    "span.break-words",
]


@dataclass(frozen=True)
class Post:
    """One collected post. ``url`` is the identity used for de-duplication."""

    url: str
    date_posted: str
    author: str
    text: str


def _first_text(container: Tag, selectors: list[str]) -> str:
    """Return cleaned text from the first matching selector, else ''."""
    for selector in selectors:
        el = container.select_one(selector)
        if el is not None:
            text = el.get_text(separator=" ", strip=True)
            if text:
                return _clean(text)
    return ""


def _clean(text: str) -> str:
    """Collapse whitespace and drop LinkedIn's duplicated a11y suffixes."""
    text = re.sub(r"\s+", " ", text).strip()
    # LinkedIn often renders "2d •" or a trailing bullet on the date line.
    text = text.strip(" •·").strip()
    return text


def _extract_urn(container: Tag) -> str:
    """Find the activity URN so we can build a stable permalink.

    Post cards carry the URN in a ``data-urn`` attribute on the container or a
    descendant, e.g. ``urn:li:activity:7123456789012345678``.
    """
    for el in [container, *container.select("[data-urn]")]:
        urn = el.get("data-urn", "") if isinstance(el, Tag) else ""
        if urn and "activity" in urn:
            return urn
    return ""


def _extract_url(container: Tag) -> str:
    """Best-effort permalink for the post.

    Prefer building one from the activity URN (stable, canonical). Fall back to
    an explicit anchor href if present.
    """
    urn = _extract_urn(container)
    if urn:
        return f"https://www.linkedin.com/feed/update/{urn}/"

    for a in container.select("a[href]"):
        href = a.get("href", "")
        if "/feed/update/" in href or "/posts/" in href:
            if href.startswith("/"):
                href = "https://www.linkedin.com" + href
            return href.split("?")[0]
    return ""


def parse_posts(html: str) -> list[Post]:
    """Parse page HTML into a list of :class:`Post`.

    Cards without a resolvable URL are skipped, since URL is our identity and a
    post we cannot link to is not useful in the spreadsheet.
    """
    soup = BeautifulSoup(html, "html.parser")

    containers: list[Tag] = []
    for selector in POST_CONTAINER_SELECTORS:
        containers.extend(soup.select(selector))

    # De-dupe containers that matched more than one selector, preserving order.
    seen_ids: set[int] = set()
    unique_containers: list[Tag] = []
    for c in containers:
        if id(c) not in seen_ids:
            seen_ids.add(id(c))
            unique_containers.append(c)

    posts: list[Post] = []
    for container in unique_containers:
        url = _extract_url(container)
        if not url:
            continue
        posts.append(
            Post(
                url=url,
                date_posted=_first_text(container, DATE_SELECTORS),
                author=_first_text(container, AUTHOR_SELECTORS),
                text=_first_text(container, TEXT_SELECTORS),
            )
        )
    return posts
