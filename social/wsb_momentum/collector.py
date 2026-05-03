"""One snapshot pass over r/wallstreetbets.

Run as:
    python -m social.wsb_momentum.collector

Schedule with Windows Task Scheduler (or cron) on a 30-minute cadence.
This script is stateless and idempotent — re-running it is safe.

Terminology: "run" / "pass" throughout this module means one execution of
collector.py — i.e. one 30-minute snapshot. NOT one day. The two invariants
that follow from this:
  - `posts.source_listing` is WRITE-ONCE on first insert. If a post is first
    seen in `hot` on run T and later appears in `top` on run T+1, the row's
    source_listing stays 'hot'.
  - `upvote_snapshots` accumulates one row per (post_id, snapshot_utc) on
    every run, regardless of which listing surfaced the post on that run.
"""

from __future__ import annotations

import time
import traceback
from datetime import datetime, timezone

from shared.reddit_json import fetch_listing, fetch_post

from . import db
from .config import (
    DB_PATH,
    LISTINGS,
    REDDIT_FETCH_LIMIT,
    REFRESH_WINDOW_HOURS,
    SLEEP_BETWEEN_LISTINGS_S,
    SLEEP_BETWEEN_POST_FETCHES_S,
    SUBREDDIT,
    TOP_TIMEFRAME,
    USER_AGENT,
)
from .ticker_extractor import extract_tickers


def _process_post(
    conn,
    post: dict,
    *,
    listing: str,
    snapshot_utc: int,
) -> tuple[bool, int]:
    """Upsert post row, ticker mappings, and this run's upvote snapshot.

    Returns (is_new_post, num_tickers_extracted).
    """
    post_id = post["id"]
    title = post.get("title") or ""
    body = post.get("selftext") or ""

    is_new = db.upsert_post(
        conn,
        post_id=post_id,
        title=title,
        author=post.get("author"),
        body=body,
        created_utc=int(post.get("created_utc") or 0),
        first_seen_utc=snapshot_utc,
        source_listing=listing,
        permalink=post.get("permalink"),
        flair=post.get("link_flair_text"),
    )

    tickers = extract_tickers(f"{title}\n{body}")
    db.add_post_tickers(conn, post_id, tickers)

    db.record_upvote_snapshot(
        conn,
        post_id=post_id,
        snapshot_utc=snapshot_utc,
        score=int(post.get("score") or 0),
        num_comments=int(post.get("num_comments") or 0),
        upvote_ratio=post.get("upvote_ratio"),
    )

    return is_new, len(tickers)


def _refresh_active_posts(conn, *, snapshot_utc: int, seen_ids: set[str]) -> tuple[int, int]:
    """Re-snapshot posts within REFRESH_WINDOW_HOURS that the listing scans missed.

    Returns (refreshed_count, gone_count).
    """
    candidates = db.get_recent_post_ids(conn, since_hours=REFRESH_WINDOW_HOURS)
    to_refresh = [pid for pid in candidates if pid not in seen_ids]
    refreshed = 0
    gone = 0

    for pid in to_refresh:
        try:
            post = fetch_post(pid, user_agent=USER_AGENT)
        except Exception as e:
            print(f"    refresh {pid} FAILED: {e!r}")
            time.sleep(SLEEP_BETWEEN_POST_FETCHES_S)
            continue

        if post is None:
            # Removed / deleted / 404 — log and skip writing a snapshot this pass.
            print(f"    refresh {pid}: post unavailable (removed/deleted), skipping")
            gone += 1
            time.sleep(SLEEP_BETWEEN_POST_FETCHES_S)
            continue

        try:
            db.record_upvote_snapshot(
                conn,
                post_id=pid,
                snapshot_utc=snapshot_utc,
                score=int(post.get("score") or 0),
                num_comments=int(post.get("num_comments") or 0),
                upvote_ratio=post.get("upvote_ratio"),
            )
            refreshed += 1
        except Exception as e:
            print(f"    refresh {pid}: write failed: {e!r}")
        time.sleep(SLEEP_BETWEEN_POST_FETCHES_S)

    if refreshed or gone:
        conn.commit()

    return refreshed, gone


def run_once() -> None:
    snapshot_utc = int(time.time())
    iso = datetime.fromtimestamp(snapshot_utc, tz=timezone.utc).isoformat()
    print(f"[snapshot @ {iso}] starting pass over r/{SUBREDDIT}")

    conn = db.init_db(DB_PATH)
    try:
        seen_ids: set[str] = set()
        new_posts = 0
        ticker_mentions = 0

        # ---- Pass 1: scan the four listings -------------------------------
        for listing in LISTINGS:
            try:
                posts = fetch_listing(
                    SUBREDDIT,
                    listing,
                    limit=REDDIT_FETCH_LIMIT,
                    timeframe=TOP_TIMEFRAME if listing == "top" else None,
                    user_agent=USER_AGENT,
                )
            except Exception as e:
                # Per spec: log and continue so the next 30-min run can pick up.
                print(f"  [{listing}] FETCH FAILED: {e!r}")
                traceback.print_exc()
                continue

            print(f"  [{listing}] fetched {len(posts)} posts")
            for post in posts:
                pid = post.get("id")
                if not pid or pid in seen_ids:
                    # If a post appears in multiple listings during the same run,
                    # we still want exactly one snapshot at this snapshot_utc.
                    continue
                seen_ids.add(pid)
                try:
                    is_new, n_tickers = _process_post(
                        conn, post, listing=listing, snapshot_utc=snapshot_utc,
                    )
                except Exception as e:
                    print(f"    post {pid} FAILED: {e!r}")
                    continue
                if is_new:
                    new_posts += 1
                ticker_mentions += n_tickers

            conn.commit()
            time.sleep(SLEEP_BETWEEN_LISTINGS_S)

        print(
            f"[snapshot @ {iso}] saw {len(seen_ids)} posts, "
            f"{new_posts} new, {ticker_mentions} ticker mentions"
        )

        # ---- Pass 2: refresh recent posts that no longer surface ----------
        refreshed, gone = _refresh_active_posts(
            conn, snapshot_utc=snapshot_utc, seen_ids=seen_ids,
        )
        print(
            f"[refresh pass] re-snapshotted {refreshed} active posts "
            f"(<{REFRESH_WINDOW_HOURS}h old), {gone} unavailable"
        )
    finally:
        conn.close()


if __name__ == "__main__":
    run_once()
