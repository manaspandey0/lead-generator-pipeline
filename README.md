# Public Data Lead Generation Pipeline

A multi-source public-data lead generation pipeline demonstrating how publicly available information can be discovered, normalized, deduplicated, enriched where explicitly available, and exported into structured datasets.

## Project Overview

This repository demonstrates three complementary public-data collection workflows:

* **NSE** — public market-member and broker records
* **SEBI** — public registered intermediary records
* **TradingQnA** — public trading-community user discovery

The project uses traders and financial-market communities as a practical demonstration of how manual prospect research can be converted into a repeatable data pipeline.

### What the project demonstrates

```text
Public Sources
      ↓
Source-specific Scrapers
      ↓
Record Discovery
      ↓
Normalization
      ↓
Deduplication
      ↓
Public Profile Enrichment
      ↓
Categorization / Keywords
      ↓
Structured Lead Data
      ↓
CSV / Excel
```

The repository contains the collection and processing code, not the scraped personal-data dataset.

---

# Data Sources

## NSE

The NSE scraper collects publicly available market-member/broker information exposed through NSE's public member resources.

The resulting records can be used for market research, organization discovery, and structured lead-generation workflows.

## SEBI

The SEBI scraper collects publicly available information from SEBI's recognized intermediary records.

The resulting dataset is intended for structured research and organization/intermediary discovery.

## TradingQnA

The TradingQnA scraper discovers individual public community members from trading-related discussions and then checks their public profile pages for information they have explicitly chosen to publish.

It is intentionally conservative:

* it reads public HTML pages only
* it checks `robots.txt` before making requests
* it does not log in
* it does not access private profiles
* it does not bypass CAPTCHA or bot protection
* it does not guess or infer contact information
* it does not search external sites for hidden contact information

---

# TradingQnA Public Profile Collector

## What it collects

For each publicly discovered person, the TradingQnA collector can produce fields including:

* `name`, `username`, `profile_url`
* `email`, `phone` — only when explicitly published on the public profile
* `whatsapp`, `telegram`, `linkedin`, `twitter`, `instagram`, `youtube` — only when the profile itself links to that account
* `website` — only when a link in the person's own bio is clearly presented as their website
* `bio` — public profile bio text
* `trading_categories`, `trading_keywords` — categories and trading-related keywords associated with the discovered topics
* `topics_seen`, `posts_seen` — counts derived from the public HTML structures that can be attributed to the person

The schema also contains:

```text
city
state
country
```

These fields remain empty when the public non-JavaScript profile page does not expose location information.

---

# How the TradingQnA Scraper Works

```text
/categories
      ↓
Target public categories
      ↓
/c/<category>/<id>
      ↓
Paginated topic lists
      ↓
/t/<slug>/<id>
      ↓
Public post authors
      ↓
/u/<username>
      ↓
Public profile HTML
      ↓
Structured person records
      ↓
CSV / XLSX
```

The scraper uses:

* `requests`
* `lxml`
* plain server-rendered HTML

It does not use TradingQnA's JSON/API endpoints and does not execute JavaScript.

Before crawling, the scraper fetches `/robots.txt` and checks whether the public paths it requires are permitted. Each subsequent request is checked again, must remain on `tradingqna.com`, and uses limited retries with backoff.

A persistent `429`, `401`, or `403` stops the run rather than attempting to work around the restriction.

---

# Installation

```bash
python -m venv .venv
```

### Windows

```powershell
.venv\Scripts\activate
```

### macOS / Linux

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

Requires Python 3.11+.

---

# Running the TradingQnA Collector

### Small test

```bash
python -m scrapers.tradingqna --max-pages 2 --max-topics 50 --max-profiles 50
```

### Larger collection

```bash
python -m scrapers.tradingqna --fresh --max-pages 30 --max-topics 12000 --max-profiles 12000
```

### Resume an existing run

```bash
python -m scrapers.tradingqna --max-pages 30 --max-topics 12000 --max-profiles 12000
```

For the full command-line options:

```bash
python -m scrapers.tradingqna --help
```

---

# Checkpoint and Resume

The scraper maintains a local cache:

```text
data/tradingqna_people_cache.json
```

Progress is periodically saved while processing categories, topics, and profiles.

CSV/XLSX outputs are checkpointed periodically during the run so a long collection has usable intermediate output.

If the process is interrupted with `Ctrl+C` or encounters an unexpected error, the cache is saved so the run can be resumed.

Use the same command again **without `--fresh`** to continue from the existing cache.

### Important output behavior

The exported CSV/XLSX reflects the topics and profiles selected for the current run.

For example, running a small test against an existing large cache can produce a smaller output dataset.

To export the complete discovered set, use limits large enough to include the full cache or use `0` where supported for no limit.

---

# Output Files

TradingQnA outputs:

```text
output/tradingqna_people.csv
output/tradingqna_people.xlsx
data/tradingqna_people_cache.json
```

The Excel workbook contains:

* `people` — one row per discovered person
* `discovery` — individually attributed public posts when the public HTML structure provides verified post attribution

Generated datasets and caches are excluded from the public repository because they may contain scraped personal information.

---

# Example Lead Generation Workflow

The project demonstrates how a company could turn public research into a repeatable prospect-discovery workflow.

For example, a company offering trading software could focus on public discussions related to:

```text
Algos
Options
Futures
Intraday trading
Technical analysis
Stocks
Nifty / Bank Nifty
```

The workflow becomes:

```text
Trading discussion
      ↓
Public author
      ↓
Public profile
      ↓
Trading categories / keywords
      ↓
Publicly published contact or social information
      ↓
Structured prospect record
```

The resulting dataset can then be reviewed and qualified before being used in a business workflow.

---

# Known Data-Source Limitations

### Public contact information is often unavailable

Many public profiles do not publish email addresses, phone numbers, websites, or social handles.

An empty contact field therefore does not necessarily indicate a collection failure.

### A username is not proof of current trading activity

A discovered username only indicates that the public account appeared in the crawled discussions or public references.

The scraper does not independently verify:

* current trading activity
* account status
* intent to buy
* professional status
* customer status

### Post attribution depends on public HTML structure

When the public HTML exposes individually attributed post blocks, the scraper can record post-level information.

When that structure is unavailable, the scraper may fall back to discovering public `/u/<username>` links. Those fallback discoveries do not create verified post records or increment `posts_seen`.

### Location availability

The current public non-JavaScript profile view does not expose a usable location field, so:

```text
city
state
country
```

may remain empty.

---

# What This Project Does Not Do

* No login or authenticated access
* No private-profile access
* No CAPTCHA bypass
* No bot-protection bypass
* No robots.txt bypass
* No stealth or browser impersonation
* No API/JSON endpoint scraping for TradingQnA
* No guessed email addresses or phone numbers
* No hidden-contact discovery on external websites

---

# Project Structure

```text
lead-generator-pipeline/
│
├── scrapers/
│   ├── nse.py
│   ├── sebi.py
│   └── tradingqna.py
│
├── config/
├── processors/
├── tests/
│
├── README.md
├── requirements.txt
├── .gitignore
└── LICENSE
```

Generated files such as:

```text
.venv/
data/
output/
logs/
__pycache__/
.env
```

are excluded from version control.

---

# Demonstration Results

A previous TradingQnA demonstration run produced:

```text
Topics collected:       12,000
Unique people:           7,575
Profiles fetched:        7,500+
```

These numbers represent a specific collection run and should not be interpreted as a permanent count of users or topics on the website.

---

# Why This Project

The goal is to demonstrate the engineering pattern behind automated public-data lead generation:

```text
Discover
   ↓
Extract
   ↓
Normalize
   ↓
Deduplicate
   ↓
Enrich
   ↓
Classify
   ↓
Export
```

The trader/community use case provides a practical example, while the same pattern can be adapted to other legitimate public-data research and prospect-discovery workflows.

---

# License

MIT — see [LICENSE](LICENSE).
