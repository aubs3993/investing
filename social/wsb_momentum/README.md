# WSB Momentum Tracker

Snapshots r/wallstreetbets posts every 30 minutes, extracts stock tickers
mentioned in those posts, and pulls matching stock prices at each snapshot
timestamp. Goal: enable later analysis of whether changes in a post's upvote
velocity correlate with (and possibly lead) changes in the underlying stock's
price.

This is the data-collection layer. The correlation analysis layer is a future
phase — do not build it yet.

## Status (May 2026)

Data collection pipeline is operational. Analysis layer is a separate future
phase — to be built once 2-4 weeks of data has accumulated. This repo
currently provides the infrastructure to collect WSB post snapshots,
intraday prices, and short-interest fundamentals; it does not yet produce
trade signals, correlation analyses, or recommendations.

## Layout

```
social/wsb_momentum/
  collector.py            # one snapshot pass over Reddit (entry point)
  price_fetcher.py        # one yfinance pass for active tickers (entry point)
  ticker_extractor.py     # regex + ticker list + blacklist
  db.py                   # SQLite schema + helpers
  config.py               # constants (UA, listings, blacklist, paths)
  data/
    tickers_nyse_nasdaq.csv  # NYSE+NASDAQ reference list (~11k rows)
  output/
    wsb.db                # created on first run; gitignored
  tests/
    test_ticker_extractor.py
    test_db.py
shared/
  reddit_json.py          # public Reddit JSON wrapper (no OAuth)
```

## Manual run

From the repo root:

```powershell
# one snapshot pass over r/wallstreetbets (every 30 min)
python -m social.wsb_momentum.collector

# fetch yfinance 15-min bars for tickers seen in the last 7 days (every 30 min)
python -m social.wsb_momentum.price_fetcher

# fetch daily fundamentals (shares outstanding, short interest, float) for active tickers (once a day)
python -m social.wsb_momentum.fundamentals_fetcher
```

Both scripts are idempotent — re-running them is safe and will not duplicate
rows. The collector locks `source_listing` on first insert (so a post seen in
`hot` and later in `top` keeps `source_listing='hot'`); upvote snapshots
accumulate on every run.

Throughout this project, **"run" / "pass" means a single execution of
`collector.py` — one 30-minute snapshot, not a day.** `source_listing` is
write-once on the first run that sees a post; `upvote_snapshots` gains one
row per post on every subsequent run regardless of which listing re-surfaced
the post.

## Setup

The repo root `requirements.txt` already lists `requests`, `yfinance`, and
`pandas`. From a fresh checkout:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

(If you skip the venv, `pip install --user` works on Windows.)

The User-Agent for Reddit is set in `config.py` and can be overridden via the
`WSB_MOMENTUM_UA` env var. Reddit's public JSON endpoints don't require auth —
do not set up OAuth or install `praw`.

## Scheduling on Windows (30-minute cadence)

The collector and the price fetcher are independent. Schedule both with
`schtasks`. Replace the path with your local clone if different.

`schtasks /create` does not directly expose the "Wake the computer to run
this task" flag, so we create the task first and then enable wake via
`PowerShell`'s `Set-ScheduledTask` cmdlet.

```powershell
# 1) collector — runs at :00 and :30 every hour, indefinitely
schtasks /create /tn "WSB Momentum - Collector" `
  /tr "powershell -NoProfile -WindowStyle Hidden -Command `"cd 'C:\Users\aubre\Documents\python_projects\investing'; python -m social.wsb_momentum.collector`"" `
  /sc minute /mo 30 /st 00:00 /f

# 2) price fetcher — same cadence, offset 5 minutes so it doesn't fight the collector
schtasks /create /tn "WSB Momentum - PriceFetcher" `
  /tr "powershell -NoProfile -WindowStyle Hidden -Command `"cd 'C:\Users\aubre\Documents\python_projects\investing'; python -m social.wsb_momentum.price_fetcher`"" `
  /sc minute /mo 30 /st 00:05 /f

# 3) fundamentals fetcher — DAILY at 5:00am ET (= 09:00 UTC; PowerShell uses local time).
#    Adjust the /st value to 05:00 if your machine is on US Eastern.
#    This runs before US market open so each trading day starts with fresh
#    shares-outstanding / short-interest / float data on file.
schtasks /create /tn "WSB Momentum - FundamentalsFetcher" `
  /tr "powershell -NoProfile -WindowStyle Hidden -Command `"cd 'C:\Users\aubre\Documents\python_projects\investing'; python -m social.wsb_momentum.fundamentals_fetcher`"" `
  /sc daily /st 05:00 /f

# 4) Enable "Wake the computer to run this task" on all three tasks
foreach ($name in @("WSB Momentum - Collector", "WSB Momentum - PriceFetcher", "WSB Momentum - FundamentalsFetcher")) {
    $task = Get-ScheduledTask -TaskName $name
    $task.Settings.WakeToRun = $true
    Set-ScheduledTask -InputObject $task | Out-Null
}
```

**Why daily for fundamentals (and not 30-min like the other two)?** The
fundamentals fields refresh on very different cadences upstream, all of
them slower than daily:

- **Short interest** (`shares_short`, `short_ratio`, `short_pct_float`) —
  published by FINRA **twice a month** with a roughly **two-week lag**.
- **Institutional ownership** (`held_pct_institutions`, and therefore
  `held_pct_retail`) — derived from **13F filings published quarterly**
  with a **~45-day lag**, so this split is always **1–3 months stale**.
  Acceptable for thesis-level analysis (the retail/institutional split
  doesn't move much quarter-over-quarter for established names) but worth
  knowing — it's not a real-time signal.
- **Float and shares outstanding** — update on corporate actions
  (buybacks, secondary offerings, splits, employee grants vesting), so
  daily is the right cadence for those.
- **Insider holdings** (`held_pct_insiders`) — updated on Form 4 filings
  within ~2 business days of a transaction, so daily here is appropriate.

A daily cadence is technically overkill for short interest and 13F-derived
ownership specifically, but harmless — it keeps scheduling consistent and
the workload trivial.

**Important — disable hibernation, leave sleep alone.** Wake-to-run works
from sleep (S3) but **does not work from hibernation (S4)**. If your machine
hibernates instead of sleeping, the scheduled tasks silently miss runs until
you wake the machine manually. Disable hibernation once with:

```powershell
# run once from an elevated PowerShell prompt
powercfg /hibernate off
```

Sleep is fine; hibernation is what breaks it.

Inspect / remove with:

```powershell
schtasks /query /tn "WSB Momentum - Collector"
schtasks /delete /tn "WSB Momentum - Collector" /f
```

If you want logs, append `>> output\collector.log 2>&1` to the inner command.

### cron equivalent (Mac/Linux)

```
*/30 * * * * cd /path/to/investing && /usr/bin/python3 -m social.wsb_momentum.collector >> social/wsb_momentum/output/collector.log 2>&1
5,35 * * * * cd /path/to/investing && /usr/bin/python3 -m social.wsb_momentum.price_fetcher >> social/wsb_momentum/output/price_fetcher.log 2>&1
0 5 * * * cd /path/to/investing && /usr/bin/python3 -m social.wsb_momentum.fundamentals_fetcher >> social/wsb_momentum/output/fundamentals.log 2>&1
```

### Limitations of scheduled collection

- **PC must be on or sleeping, not off.** Wake-to-run pulls the machine out
  of S3 sleep but can't power it on from a fully shut-down state. If the PC
  is off when a 30-min slot fires, that snapshot is missed.
- **Hibernation breaks it.** See above. Run `powercfg /hibernate off` once.
- **Gaps are detectable in analysis.** Look for discontinuities in
  `snapshot_utc` — i.e., adjacent rows where the gap between
  `snapshot_utc` values is materially larger than 1800 seconds. A reliable
  detection pattern is `LAG(snapshot_utc) OVER (PARTITION BY post_id
  ORDER BY snapshot_utc)` with a `> 2700` (45-min) threshold. The
  correlation analysis layer should treat these gaps as missing
  observations rather than collapsing them into adjacent points.

## Tests

```powershell
pytest social/wsb_momentum/tests/
```

## Auditing the ticker extractor (manual tuning)

`audit_tickers.py` is a read-only utility for finding false-positive ticker
extractions. It never touches the DB or `config.py` — it just prints a ranked
report you can use to decide which entries to add to `TICKER_BLACKLIST`
manually.

```powershell
# default: rank every ticker mentioned in the last 7 days by suspicion score
python -m social.wsb_momentum.audit_tickers

# deep-dive: show every mention of one ticker with surrounding text context
python -m social.wsb_momentum.audit_tickers --ticker GME
```

The default report ends with a "Suggested blacklist additions" block; copy
the lines you agree with into `TICKER_BLACKLIST` in `config.py`. Example
real-world finding from this project: `P` showed up at the top of mention
counts because of `P/E`, `P/S`, `P/FFO`, `S&P`, and `P&L` — confirmable in
under 10 seconds via `--ticker P`.

## Smoke-test SQL queries

Open `social/wsb_momentum/output/wsb.db` with the `sqlite3` CLI or any viewer.
These four queries sanity-check the data shape; they are also the analytical
building blocks for the future correlation layer.

### 1. Top 10 tickers by mention count in last 24h

Sanity check on volume — does the pipeline see tickers at the rate you'd expect?

```sql
SELECT
  pt.ticker,
  COUNT(DISTINCT pt.post_id) AS posts_mentioning,
  COUNT(*) AS total_mentions
FROM post_tickers pt
JOIN posts p ON p.id = pt.post_id
WHERE p.first_seen_utc >= strftime('%s', 'now', '-1 day')
GROUP BY pt.ticker
ORDER BY total_mentions DESC
LIMIT 10;
```

### 2. A specific post's upvote trajectory over time, joined to ticker

Sanity check on the time series — pick a post id from query 1's results.

```sql
SELECT
  us.snapshot_utc,
  datetime(us.snapshot_utc, 'unixepoch') AS snapshot_iso,
  us.score,
  us.num_comments,
  us.upvote_ratio,
  GROUP_CONCAT(pt.ticker) AS tickers
FROM upvote_snapshots us
LEFT JOIN post_tickers pt ON pt.post_id = us.post_id
WHERE us.post_id = '1t225f3'  -- replace with a real post id
GROUP BY us.snapshot_utc
ORDER BY us.snapshot_utc;
```

### 3. Ticker-level aggregate over time

Sum of scores and active post count across all posts mentioning a given ticker,
grouped by snapshot timestamp. This is the "attention" series for a ticker.

```sql
SELECT
  pt.ticker,
  us.snapshot_utc,
  SUM(us.score) AS total_score,
  COUNT(DISTINCT us.post_id) AS active_posts,
  MAX(us.score) AS max_post_score
FROM upvote_snapshots us
JOIN post_tickers pt ON pt.post_id = us.post_id
WHERE pt.ticker = 'GME'
GROUP BY pt.ticker, us.snapshot_utc
ORDER BY us.snapshot_utc;
```

### 4. Ticker-level momentum (first difference of total score) — the future correlation foundation

Window function over query 3 to compute upvotes-per-30-min — the actual
momentum signal. **This `momentum_30m` series is what we will eventually
correlate against the price-return series in `price_snapshots` (with various
lags) to ask whether WSB momentum leads price.** The analysis layer that joins
this to price returns is intentionally not built yet.

```sql
WITH ticker_totals AS (
  SELECT
    pt.ticker,
    us.snapshot_utc,
    SUM(us.score) AS total_score,
    COUNT(DISTINCT us.post_id) AS active_posts
  FROM upvote_snapshots us
  JOIN post_tickers pt ON pt.post_id = us.post_id
  WHERE pt.ticker = 'GME'
  GROUP BY pt.ticker, us.snapshot_utc
)
SELECT
  ticker,
  snapshot_utc,
  total_score,
  active_posts,
  total_score - LAG(total_score) OVER (PARTITION BY ticker ORDER BY snapshot_utc) AS momentum_30m
FROM ticker_totals
ORDER BY snapshot_utc;
```

### 5. Squeeze candidates — high WSB momentum on a structurally squeezable setup

Joins recent WSB momentum to the latest fundamentals row per ticker and
filters to setups where (a) shorts are sized into a constrained float and
(b) the public float is itself a small fraction of total shares
outstanding.

**Why this query matters.** Stocks with high short interest as % of float
*and* a small, restricted float are structurally more squeezable. Shorts
covering creates forced buying pressure, and limited tradeable supply can't
easily absorb it. WSB momentum alone is a weak signal; WSB momentum on a
setup like this is the regime where price moves get violent.

The two computed ratios are the squeeze-relevant metrics:
- `short_pct_float` — high means lots of forced buyers when sentiment turns.
- `float_pct_outstanding` — measures how much of the company's stock is
  actually tradeable versus held by insiders or restricted. Low is more
  squeezable because the effective supply is constrained.

This query is a building block for the future correlation analysis layer,
where we'll test whether WSB momentum has stronger predictive power on
price for tickers in this subset versus the broader universe.

```sql
WITH recent_momentum AS (
  SELECT
    pt.ticker,
    SUM(us.score) AS total_score_24h,
    COUNT(DISTINCT us.post_id) AS active_posts_24h
  FROM upvote_snapshots us
  JOIN post_tickers pt ON pt.post_id = us.post_id
  WHERE us.snapshot_utc > strftime('%s', 'now', '-1 day')
  GROUP BY pt.ticker
),
latest_fundamentals AS (
  SELECT ticker, shares_short, float_shares, short_pct_float, short_ratio,
         float_pct_outstanding, held_pct_institutions, held_pct_insiders, held_pct_retail
  FROM ticker_fundamentals tf1
  WHERE snapshot_date = (
    SELECT MAX(snapshot_date) FROM ticker_fundamentals tf2 WHERE tf2.ticker = tf1.ticker
  )
)
SELECT
  rm.ticker,
  rm.total_score_24h,
  rm.active_posts_24h,
  lf.short_pct_float,
  lf.short_ratio AS days_to_cover,
  lf.float_shares,
  lf.float_pct_outstanding,
  lf.held_pct_retail,
  lf.held_pct_institutions
FROM recent_momentum rm
JOIN latest_fundamentals lf ON lf.ticker = rm.ticker
WHERE lf.short_pct_float > 0.20         -- >20% of float is short
  AND lf.float_shares < 100000000        -- <100M share float (small enough to squeeze)
  AND lf.float_pct_outstanding < 0.80    -- <80% of shares are public float
  AND lf.held_pct_retail > 0.30          -- meaningful retail ownership (>30%)
ORDER BY rm.total_score_24h DESC;
```

### 6. Retail-attention candidates — high WSB activity on retail-dominant names

Identifies the GME/AMC pattern (retail-dominant ownership + sustained WSB
attention) regardless of whether a squeeze setup is also present. Uses a
7-day momentum window because retail-driven runs build over days, not
hours, and uses a `>1000` cumulative-score floor to filter out noise.

**Why this matters distinct from query #5.** The squeeze query (#5) is
mechanism-first — it surfaces names where shorts are forced to buy. This
query is *coalition*-first — it surfaces names where the retail crowd
already owns the float. Both setups can produce sharp WSB-led moves, but
the price dynamics differ: squeezes are typically violent and short
(forced covering), retail-dominant runs are slower and more reflexive
(coordination effects). The future correlation analysis layer will
likely want to test these subsets separately.

```sql
WITH recent_momentum AS (
  SELECT
    pt.ticker,
    SUM(us.score) AS total_score_7d,
    COUNT(DISTINCT us.post_id) AS active_posts_7d
  FROM upvote_snapshots us
  JOIN post_tickers pt ON pt.post_id = us.post_id
  WHERE us.snapshot_utc > strftime('%s', 'now', '-7 days')
  GROUP BY pt.ticker
),
latest_fundamentals AS (
  SELECT ticker, held_pct_retail, held_pct_institutions, float_shares, short_pct_float
  FROM ticker_fundamentals tf1
  WHERE snapshot_date = (
    SELECT MAX(snapshot_date) FROM ticker_fundamentals tf2 WHERE tf2.ticker = tf1.ticker
  )
)
SELECT
  rm.ticker,
  rm.total_score_7d,
  rm.active_posts_7d,
  lf.held_pct_retail,
  lf.held_pct_institutions,
  lf.short_pct_float
FROM recent_momentum rm
JOIN latest_fundamentals lf ON lf.ticker = rm.ticker
WHERE lf.held_pct_retail > 0.40        -- retail-dominant (>40%)
  AND rm.total_score_7d > 1000          -- non-trivial WSB attention
ORDER BY rm.total_score_7d DESC;
```

## Snapshot lifecycle (when posts stop accumulating snapshots)

Each collector run has two passes:

1. **Listing scan** — fetch `hot`, `new`, `rising`, `top` (up to 100 each).
   Every post returned gets a snapshot at this run's timestamp.
2. **Refresh pass** — for every post in the DB whose `created_utc` is within
   the last 72 hours and which the listing scan did NOT just snapshot,
   fetch the post directly via its `/comments/{id}.json` endpoint and write
   another snapshot. This catches posts that have aged out of `hot` / `new`
   while still actively gaining or losing upvotes.

Posts are tracked actively for their **first 72 hours of life**. After that,
the row remains in the DB as a historical record but no new snapshots
accumulate. This is intentional — Reddit post scores effectively freeze after
~3 days, so additional polls would just write identical rows.

If a post is removed (deleted by user, mod-removed, quarantined, 404), the
refresh pass logs a warning and skips it for that run. Earlier snapshots are
left in place; the time series simply ends.

## Known limitations / tuning notes

- **Public JSON endpoints, no OAuth.** Reddit throttles unauthenticated
  traffic at ~60 req/min. We make 4 requests per snapshot and sleep 1s between
  listings, so we're well under the limit. If you start seeing 429s anyway,
  raise the User-Agent uniqueness or move to OAuth (a future change).
- **Universe is NYSE + NASDAQ only.** OTC, foreign listings, and crypto are
  dropped at extraction. Expand the reference CSV if you want broader coverage.
- **`S&P` produces false-positive `S` and `P` mentions** because `&` is a
  word boundary. If those bubble to the top of mention counts, add `S` and `P`
  to `TICKER_BLACKLIST` in `config.py`.
- **yfinance is best-effort.** Empty bars (delisted, SPAC, illiquid) are
  logged and skipped. 15-minute granularity is good enough for v1; tighten if
  you need finer resolution.
- **No backfill.** `first_seen_utc` is when our collector first saw a post;
  there is no upvote history for the period before we started watching.
