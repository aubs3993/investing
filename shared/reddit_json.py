"""Thin wrapper around requests for Reddit's public JSON endpoints.

No OAuth — uses the unauthenticated `/{listing}.json?limit=N` endpoints.
Reddit throttles unauthenticated traffic at roughly 60 req/min, so callers
should pace themselves. We retry once on 429 with exponential backoff.
"""

from __future__ import annotations

import time
from typing import Any

import requests

DEFAULT_UA = "wsb-momentum-research/0.1 (by u/aubs3993)"
BASE_URL = "https://www.reddit.com"
TIMEOUT_S = 15


def fetch_listing(
    subreddit: str,
    listing: str,
    limit: int = 100,
    *,
    timeframe: str | None = None,
    user_agent: str = DEFAULT_UA,
    session: requests.Session | None = None,
) -> list[dict[str, Any]]:
    """Fetch one listing of posts from a subreddit's public JSON endpoint.

    Returns the list of post `data` dicts (the inner `data.children[*].data`
    slice). Raises requests.HTTPError on non-recoverable HTTP failures.
    """
    url = f"{BASE_URL}/r/{subreddit}/{listing}.json"
    params: dict[str, Any] = {"limit": limit, "raw_json": 1}
    if timeframe and listing == "top":
        params["t"] = timeframe

    headers = {"User-Agent": user_agent}
    s = session or requests

    for attempt in range(2):
        resp = s.get(url, params=params, headers=headers, timeout=TIMEOUT_S)
        if resp.status_code == 429 and attempt == 0:
            backoff = 5.0
            time.sleep(backoff)
            continue
        resp.raise_for_status()
        payload = resp.json()
        children = payload.get("data", {}).get("children", [])
        return [c["data"] for c in children if c.get("kind") == "t3"]

    # Should not reach here — the loop either returns or raises.
    raise RuntimeError("fetch_listing exhausted retries without returning")


def fetch_post(
    post_id: str,
    *,
    user_agent: str = DEFAULT_UA,
    session: requests.Session | None = None,
) -> dict[str, Any] | None:
    """Fetch one post's current state by id via the public comments JSON.

    `post_id` may be either the bare id ('xxxxxx') or the t3-prefixed form
    ('t3_xxxxxx'); the prefix is stripped if present.

    Returns the post `data` dict, or None if the post was removed/deleted/empty
    (the response shape Reddit returns for those cases is `[{listing-of-zero},
    {comments}]` or carries `removed_by_category`/`selftext == '[removed]'`).
    """
    bare = post_id[3:] if post_id.startswith("t3_") else post_id
    url = f"{BASE_URL}/comments/{bare}.json"
    params = {"limit": 1, "raw_json": 1}
    headers = {"User-Agent": user_agent}
    s = session or requests

    for attempt in range(2):
        resp = s.get(url, params=params, headers=headers, timeout=TIMEOUT_S)
        if resp.status_code == 429 and attempt == 0:
            time.sleep(5.0)
            continue
        if resp.status_code in (403, 404):
            # Post deleted/quarantined/private — treat as gone.
            return None
        resp.raise_for_status()

        payload = resp.json()
        # Comments endpoint returns [post_listing, comments_listing].
        if not isinstance(payload, list) or not payload:
            return None
        children = payload[0].get("data", {}).get("children", [])
        if not children:
            return None
        data = children[0].get("data") or {}
        if not data.get("id"):
            return None
        # Treat mod-removed / user-deleted posts as gone for snapshot purposes —
        # their score is frozen and selftext is "[removed]" / "[deleted]".
        if data.get("removed_by_category") or data.get("selftext") in ("[removed]", "[deleted]"):
            return None
        return data

    raise RuntimeError("fetch_post exhausted retries without returning")
