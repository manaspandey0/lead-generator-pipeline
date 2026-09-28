# Public Data Lead Generation Pipeline

## Project Overview

This repository demonstrates a multi-source public-data lead generation pipeline using:

- **NSE** — public market-member/broker records
- **SEBI** — public registered intermediary records
- **TradingQnA** — public trading-community user discovery

The pipeline demonstrates how public information can be collected, normalized, deduplicated, enriched where explicitly available, and exported into structured datasets for research and prospect discovery.

A small, HTML-only research tool that discovers public community members on
[TradingQnA](https://tradingqna.com) (Zerodha's trading community forum) and
collects whatever contact/social information they have chosen to publish on
their own public profile page.

It is intentionally conservative: it only reads public HTML pages a browser
(or a search engine) could read anyway, it checks `robots.txt` before making
any request, and it never logs in, bypasses access controls, or guesses at
information a person hasn't published themselves.

## What it collects

For each publicly discovered person, one row with:

- `name`, `username`, `profile_url`
- `email`, `phone` -- only when explicitly published on their profile
- `whatsapp`, `telegram`, `linkedin`, `twitter`, `instagram`, `youtube` -- only
  when the profile itself links to that account
- `website` -- only when a link in the person's own bio is clearly labeled as
  their site ("Website:", "my blog", "portfolio", etc.); an unlabeled link
  (e.g. a cited article) is never treated as someone's personal website
- `bio` -- their public bio text, verbatim
- `trading_categories`, `trading_keywords` -- which target categories/topics
  they were seen posting in, and which trading-related keywords appeared in
  those topic titles/categories
- `topics_seen`, `posts_seen` -- how many of the crawled topics/posts this
  person appears in (see "Known data-source limitations" below)

`city`, `state`, and `country` columns are also present for schema stability,
but are always empty: TradingQnA's public (non-JavaScript) profile page does
not expose a location field at all, so there is nothing to collect there.

## How it works

```
/categories (public HTML)
      -> target categories (Trading, F&O, Stocks, ...)
      -> /c/<category>/<id> topic-list HTML, paginated
      -> topic URLs
      -> /t/<slug>/<id> topic HTML
      -> public post authors (usernames)
      -> /u/<username> public profile HTML
      -> CSV / XLSX
```

Every step reads plain server-rendered HTML with `requests` + `lxml`. There is
no use of TradingQnA's JSON/API endpoints, no JavaScript execution, and no
attempt to render or query anything other than the public page a logged-out
visitor already sees.

Before the first request, the tool fetches `/robots.txt` and refuses to run
at all if any of the public paths it needs (`/categories`, `/latest`, `/c/`,
`/t/`, `/u/`) are disallowed for `*`. Every subsequent request is checked
against the same rules, must stay on `tradingqna.com`, and is retried a
limited number of times with backoff; a persistent HTTP 429 or 401/403 stops
the run rather than working around it.

## Installation

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows
# source .venv/bin/activate   # macOS/Linux

pip install -r requirements.txt
```

Requires Python 3.11+.

## Running it

```bash
# Small test run -- fast, safe to run anytime
python -m scrapers.tradingqna --max-pages 2 --max-topics 50 --max-profiles 50

# Larger collection run
python -m scrapers.tradingqna --fresh --max-pages 30 --max-topics 7500 --max-profiles 12000

# Resume a previous run using the local cache
python -m scrapers.tradingqna --max-pages 30 --max-topics 7500 --max-profiles 12000
```

See `python -m scrapers.tradingqna --help` for the full option list.

### Resume / checkpoint behavior

Every category page, topic page, and profile page fetched is written to
`data/tradingqna_people_cache.json` as it's collected. Re-running the same
command (without `--fresh`) picks up where it left off instead of
re-requesting pages it already has. The CSV/XLSX outputs are also re-written
periodically during a run (every 100 topics/profiles) and on exit, including
after an interruption (Ctrl+C) or an unexpected error, so a long run never
loses progress.

**Caveat:** each run's exported CSV/XLSX reflects only the topics/profiles
selected for *that* run (via `--max-topics` / `--max-profiles`), not
everything ever discovered across all previous runs. Running a small test
(e.g. `--max-topics 10`) against a cache that already holds thousands of
topics will overwrite the output files with that smaller subset. If you want
the full dataset re-exported, use limits at least as large as your cache
already contains (or set them to `0` for no cap).

## Output files

- `output/tradingqna_people.csv` -- one row per discovered person
- `output/tradingqna_people.xlsx` -- same data, plus a `discovery` sheet
  (one row per individually-attributed public post -- see below)
- `data/tradingqna_people_cache.json` -- the full local cache/checkpoint;
  safe to delete to force a clean re-crawl, or keep to resume

None of these are committed to this repository (see `.gitignore`) since they
contain scraped, potentially personal data.

## Known data-source limitations

- **Not every profile publishes contact or social information.** Most public
  TradingQnA profiles have no email, phone, website, or social links at all --
  `email`/`phone`/`website`/social columns being empty is the expected,
  common case, not a collection failure.
- **A discovered username is not a guarantee of an active trader.** This tool
  discovers anyone whose public username appears as a post author (or is
  `@mentioned`) in the target categories; it does not verify trading activity,
  account status, or intent.
- **`posts_seen` / the `discovery` sheet reflect individually-attributed
  posts only.** On topic pages where TradingQnA's public HTML exposes a
  distinct post block per author (the common case), each post is counted
  individually. On the rare page where that structure isn't present, the
  tool falls back to scanning for public `/u/<username>` links so the person
  is still discovered -- but that fallback does not produce a verified post
  count, so it does not increment `posts_seen` or add a `discovery` row.
  `topics_seen` (how many distinct topics a person appears in) is unaffected
  and always accurate.
- **Location is never populated.** TradingQnA's public, non-JavaScript
  profile view has no location field to read.

## What this tool does not do

- No login, session cookies, or private-profile access
- No CAPTCHA or bot-protection bypass
- No stealth headers or browser impersonation -- requests identify themselves
  plainly via `User-Agent: TradingQnA-Public-Research/1.0`
- No use of TradingQnA's JSON/API endpoints
- No inferring, guessing, or looking up contact details anywhere off-profile

## License

MIT -- see [LICENSE](LICENSE).
