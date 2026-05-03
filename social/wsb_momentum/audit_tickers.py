"""Manual tuning utility — audit ticker extractions for false positives.

This is a READ-ONLY tool. It never writes to the DB, never edits config.py,
and never auto-applies a blacklist. It prints a ranked report to stdout so
you can decide which tickers to add to TICKER_BLACKLIST in config.py.

Usage:
    python -m social.wsb_momentum.audit_tickers
    python -m social.wsb_momentum.audit_tickers --ticker GME

Default mode: rank every ticker mentioned in the last 7 days by suspicion
score (most suspicious first), with per-ticker stats and 5 sample post titles.
A "Suggested blacklist additions" diff is printed at the bottom — copy/paste
manually if you want to apply.

Single-ticker mode (--ticker XYZ): show the full evidence trail for one
ticker — every post that mentions it, every snapshot of every post, the
text context of each mention.
"""

from __future__ import annotations

import argparse
import re
import sqlite3
import statistics
import sys
import textwrap
import time
from pathlib import Path

from .config import DB_PATH, TICKER_BLACKLIST

WINDOW_DAYS = 7
SAMPLE_TITLES = 5

# When checking "this ticker is in the blacklist but somehow still being
# extracted" (a real bug indicator), only count extractions seen within this
# many hours. Older rows are historical leftovers from before the blacklist
# entry was added — not bugs.
BLACKLIST_BUG_CUTOFF_HOURS = 24

# Common English words / pronouns / verbs / connectives that frequently look
# like tickers when capitalized. Hits on this list inflate the suspicion score.
COMMON_ENGLISH = frozenset({
    "A", "AM", "AN", "AND", "ANY", "ARE", "AS", "AT", "BE", "BUT", "BY", "DO",
    "FOR", "GO", "HAD", "HAS", "HE", "HER", "HIS", "HOW", "I", "IF", "IN",
    "IS", "IT", "ITS", "ME", "MY", "NEW", "NO", "NOT", "NOW", "OF", "ON",
    "ONE", "OR", "OUR", "OUT", "PER", "SAW", "SHE", "SO", "THE", "THIS",
    "TO", "TWO", "UP", "US", "WE", "WHO", "WHY", "YES", "YOU", "YOUR",
    "ALL", "BAD", "BIG", "CAN", "DID", "DUE", "END", "FEW", "GOT", "GOOD",
    "HOT", "JOB", "LET", "LOT", "LOW", "MAD", "MAN", "MAY", "OLD", "ONE",
    "OWN", "PAY", "PUT", "RAN", "RED", "RUN", "SAY", "SET", "SIT", "SIX",
    "TAX", "TEN", "TIP", "TOP", "TRY", "TWO", "USE", "WAR", "WAY", "WIN",
    "YET", "AGO", "WAS", "WERE", "WILL", "BEEN", "MORE", "MOST", "SOME",
    "SUCH", "ONLY", "VERY", "WHAT", "WHEN", "WHERE", "WHICH", "WITH", "FROM",
    "HAVE", "JUST", "LIKE", "MAKE", "MANY", "OVER", "TAKE", "THAN", "THEM",
    "THEY", "THIS", "TIME", "WANT", "WELL", "BACK", "DOWN", "EACH", "EVEN",
    "MUCH", "NEXT", "ONCE", "STILL",
    # Months/days short forms
    "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT",
    "NOV", "DEC", "MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN",
})

CASHTAG_RE = re.compile(r"\$([A-Z]{1,5})\b")
BARE_RE = re.compile(r"(?<![A-Za-z0-9$])([A-Z]{1,5})(?![A-Za-z0-9])")


def _connect(db_path: Path) -> sqlite3.Connection:
    if not db_path.exists():
        sys.exit(f"DB not found: {db_path}\nRun the collector first.")
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    return conn


def _suspicion_score(*, ticker: str, mentions: int, distinct_posts: int,
                     mean_score: float, median_post_len: int,
                     cashtag_ratio: float, recent_mentions: int) -> tuple[float, list[str]]:
    """Higher = more likely to be a false positive. Returns (score, reasons).

    `recent_mentions` is the count of mentions on posts first seen within
    BLACKLIST_BUG_CUTOFF_HOURS — used to gate the blacklist-bug bonus so it
    only fires for genuinely-new extractions, not historical leftovers from
    before a blacklist entry was added.
    """
    score = 0.0
    reasons: list[str] = []

    if ticker in COMMON_ENGLISH:
        score += 5.0
        reasons.append("common English word")
    if len(ticker) == 1:
        score += 2.5
        reasons.append("single letter (high collision risk)")
    elif len(ticker) == 2:
        score += 1.0
        reasons.append("two letters (collision risk)")

    # Cashtags express explicit author intent. A ticker that is ALWAYS bare and
    # NEVER cashtagged is more suspect than one with at least some cashtags.
    if mentions >= 3 and cashtag_ratio == 0.0:
        score += 1.5
        reasons.append("0% cashtag use")

    # Real ticker discussion tends to land on posts with non-trivial scores.
    # Pure noise (e.g. extracted from random rants) shows up in many low-score posts.
    if distinct_posts >= 4 and mean_score < 50:
        score += 1.0
        reasons.append(f"low mean post score ({mean_score:.0f}) across {distinct_posts} posts")

    # Very short posts capitalizing common words are often the source of noise.
    if median_post_len < 200:
        score += 0.5
        reasons.append(f"median post body length only {median_post_len} chars")

    # Already-blacklisted tickers shouldn't appear in NEW extractions. Old rows
    # in post_tickers from before the blacklist entry was added are not bugs —
    # they're historical leftovers. Only flag if extractions are still showing
    # up on posts first seen within the cutoff window.
    if ticker in TICKER_BLACKLIST and recent_mentions > 0:
        score += 10.0
        reasons.append(f"ALREADY in TICKER_BLACKLIST but extracted on {recent_mentions} "
                       f"post(s) in last {BLACKLIST_BUG_CUTOFF_HOURS}h (bug?)")

    return score, reasons


def _classify_mentions(text: str, ticker: str) -> tuple[int, int]:
    """Count cashtag vs bare occurrences of `ticker` in `text`."""
    if not text:
        return 0, 0
    cashtags = sum(1 for m in CASHTAG_RE.findall(text) if m == ticker)
    bare = sum(1 for m in BARE_RE.findall(text) if m == ticker)
    return cashtags, bare


def _gather_per_ticker_stats(conn: sqlite3.Connection) -> list[dict]:
    """Compute per-ticker stats over the last WINDOW_DAYS days."""
    rows = conn.execute(
        f"""
        SELECT pt.ticker, p.id AS post_id, p.title,
               COALESCE(p.body, '') AS body, p.first_seen_utc
        FROM post_tickers pt
        JOIN posts p ON p.id = pt.post_id
        WHERE p.first_seen_utc >= strftime('%s', 'now', '-{WINDOW_DAYS} days')
        """
    ).fetchall()

    bug_cutoff_utc = int(time.time()) - BLACKLIST_BUG_CUTOFF_HOURS * 3600

    by_ticker: dict[str, dict] = {}
    for r in rows:
        t = r["ticker"]
        d = by_ticker.setdefault(t, {
            "ticker": t, "posts": [], "cashtag_count": 0, "bare_count": 0,
            "recent_mentions": 0,
        })
        text = f"{r['title']}\n{r['body']}"
        c, b = _classify_mentions(text, t)
        d["cashtag_count"] += c
        d["bare_count"] += b
        d["posts"].append({"id": r["post_id"], "title": r["title"], "body_len": len(r["body"])})
        if r["first_seen_utc"] >= bug_cutoff_utc:
            d["recent_mentions"] += (c + b)

    # Pull max-score per post for the post-quality stats.
    post_scores = {
        r["post_id"]: r["max_score"]
        for r in conn.execute("""
            SELECT post_id, MAX(score) AS max_score
            FROM upvote_snapshots
            GROUP BY post_id
        """).fetchall()
    }

    out: list[dict] = []
    for t, d in by_ticker.items():
        scores = [post_scores.get(p["id"], 0) for p in d["posts"]]
        body_lens = [p["body_len"] for p in d["posts"]]
        mentions = d["cashtag_count"] + d["bare_count"]
        cashtag_ratio = (d["cashtag_count"] / mentions) if mentions else 0.0
        mean_score = statistics.mean(scores) if scores else 0
        median_score = statistics.median(scores) if scores else 0
        median_body = int(statistics.median(body_lens)) if body_lens else 0
        susp, reasons = _suspicion_score(
            ticker=t,
            mentions=mentions,
            distinct_posts=len(d["posts"]),
            mean_score=mean_score,
            median_post_len=median_body,
            cashtag_ratio=cashtag_ratio,
            recent_mentions=d["recent_mentions"],
        )
        out.append({
            "ticker": t,
            "mentions": mentions,
            "distinct_posts": len(d["posts"]),
            "cashtag_count": d["cashtag_count"],
            "bare_count": d["bare_count"],
            "cashtag_ratio": cashtag_ratio,
            "mean_score": mean_score,
            "median_score": median_score,
            "median_body_len": median_body,
            "suspicion": susp,
            "reasons": reasons,
            "sample_titles": [p["title"] for p in d["posts"][:SAMPLE_TITLES]],
        })

    out.sort(key=lambda x: (-x["suspicion"], -x["mentions"], x["ticker"]))
    return out


def _print_report(stats: list[dict]) -> None:
    print(f"\nWSB ticker audit — last {WINDOW_DAYS} days, {len(stats)} unique tickers")
    print("=" * 90)
    header = f"{'ticker':<6} {'susp':>5} {'ment':>5} {'posts':>5} {'cash%':>6} {'mean$':>7} {'medlen':>6}  reasons"
    print(header)
    print("-" * 90)
    for s in stats:
        reason = "; ".join(s["reasons"]) if s["reasons"] else ""
        print(f"{s['ticker']:<6} {s['suspicion']:>5.1f} "
              f"{s['mentions']:>5} {s['distinct_posts']:>5} "
              f"{s['cashtag_ratio']*100:>5.0f}% {s['mean_score']:>7.0f} "
              f"{s['median_body_len']:>6}  {reason}")

    print()
    print("Sample titles for top-15 most-suspicious tickers:")
    print("-" * 90)
    for s in stats[:15]:
        print(f"\n[{s['ticker']}]  suspicion={s['suspicion']:.1f}  "
              f"({s['mentions']} mentions in {s['distinct_posts']} posts)")
        for t in s["sample_titles"]:
            safe = t.encode("ascii", "replace").decode("ascii")
            print(f"  - {safe[:100]}")

    suggestions = [s["ticker"] for s in stats if s["suspicion"] >= 4.0
                   and s["ticker"] not in TICKER_BLACKLIST]
    print()
    print("=" * 90)
    print("Suggested blacklist additions (suspicion >= 4.0, not yet blacklisted):")
    print("Manual review required — copy lines you agree with into TICKER_BLACKLIST")
    print("in social/wsb_momentum/config.py. Nothing is auto-applied.")
    print("-" * 90)
    if not suggestions:
        print("(no suggestions — extraction looks clean over this window)")
    else:
        for t in suggestions:
            print(f'    "{t}",')


def _print_single_ticker(conn: sqlite3.Connection, ticker: str) -> None:
    ticker = ticker.upper()
    print(f"\nDeep-dive: ticker '{ticker}' (last {WINDOW_DAYS} days)")
    print("=" * 90)

    posts = conn.execute(
        f"""
        SELECT p.id, p.title, p.body, p.author, p.source_listing,
               p.created_utc, p.first_seen_utc, p.permalink,
               (SELECT MAX(score) FROM upvote_snapshots us WHERE us.post_id = p.id)        AS max_score,
               (SELECT COUNT(*)   FROM upvote_snapshots us WHERE us.post_id = p.id)        AS n_snapshots
        FROM post_tickers pt
        JOIN posts p ON p.id = pt.post_id
        WHERE pt.ticker = ?
          AND p.first_seen_utc >= strftime('%s', 'now', '-{WINDOW_DAYS} days')
        ORDER BY max_score DESC NULLS LAST
        """,
        (ticker,),
    ).fetchall()

    if not posts:
        print(f"No posts mention {ticker} in the last {WINDOW_DAYS} days.")
        return

    if ticker in TICKER_BLACKLIST:
        print(f"NOTE: {ticker} is currently in TICKER_BLACKLIST.")

    print(f"{len(posts)} post(s) mention {ticker}\n")

    for r in posts:
        text = f"{r['title']}\n{r['body'] or ''}"
        cashtags, bare = _classify_mentions(text, ticker)
        safe_title = r["title"].encode("ascii", "replace").decode("ascii")
        print(f"[{r['id']}]  src={r['source_listing']}  author={r['author']}  "
              f"max_score={r['max_score']}  snaps={r['n_snapshots']}  "
              f"cashtag={cashtags}  bare={bare}")
        print(f"  title: {safe_title[:120]}")
        # Show the actual mention contexts.
        for kind, regex in (("cashtag", CASHTAG_RE), ("bare", BARE_RE)):
            for m in regex.finditer(text):
                if m.group(1) != ticker:
                    continue
                start = max(0, m.start() - 40)
                end = min(len(text), m.end() + 40)
                ctx = text[start:end].replace("\n", " ")
                ctx = ctx.encode("ascii", "replace").decode("ascii")
                print(f"  {kind}: ...{ctx}...")
        print()


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument(
        "--ticker", "-t",
        help="Deep-dive a single ticker instead of the ranked report.",
    )
    parser.add_argument(
        "--db",
        default=str(DB_PATH),
        help=f"Path to the wsb.db file (default: {DB_PATH})",
    )
    args = parser.parse_args(argv)

    conn = _connect(Path(args.db))
    try:
        if args.ticker:
            _print_single_ticker(conn, args.ticker)
        else:
            stats = _gather_per_ticker_stats(conn)
            _print_report(stats)
    finally:
        conn.close()


if __name__ == "__main__":
    main()
