"""
scrapers/tradingqna.py

PUBLIC, HTML-ONLY TradingQnA person discovery collector.

Purpose
-------
Discover individual public community authors from TradingQnA's public HTML pages,
then visit public HTML profile pages and collect only information that is visibly
published on that profile.

This script intentionally does NOT use TradingQnA JSON/API endpoints.

It does NOT:
- log in
- access private profiles
- bypass CAPTCHA / bot protection / rate limits
- guess or infer contact details
- search external sites for a person's hidden contact details
- scrape arbitrary post text for contact details

Public profile contacts are collected only when the profile itself visibly
publishes them (mailto/tel links, or contact text in the user's own bio).

Notes on data availability
---------------------------
TradingQnA (a Discourse forum) serves a lightweight "crawler-safe" HTML view to
non-JS clients such as this scraper. That view exposes only: avatar, username,
and the user's own bio text. It does not include a dedicated website/location
field, so `city`, `state`, and `country` are always empty for this source, and
`website`/social fields are only populated when the user's bio itself contains
an unambiguous link (see `extract_profile_links`).

First test:
    python -m scrapers.tradingqna --max-pages 2 --max-topics 50 --max-profiles 50

Scale run:
    python -m scrapers.tradingqna --fresh --max-pages 30 --max-topics 7500 --max-profiles 12000

Resume:
    python -m scrapers.tradingqna --max-pages 30 --max-topics 7500 --max-profiles 12000

Fresh restart:
    python -m scrapers.tradingqna --fresh --max-pages 2 --max-topics 50

Outputs:
    output/tradingqna_people.csv
    output/tradingqna_people.xlsx
    data/tradingqna_people_cache.json
"""

from __future__ import annotations

import argparse
import hashlib
import html as html_lib
import json
import re
import sys
import time
from pathlib import Path
from urllib import robotparser
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit

import pandas as pd
import requests
from lxml import html as lxml_html
from lxml.html import HtmlElement

# --------------------------------------------------------------------------- config

BASE_URL = "https://tradingqna.com"
CATEGORIES_URL = f"{BASE_URL}/categories"

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"

CACHE_PATH = DATA_DIR / "tradingqna_people_cache.json"
PEOPLE_CSV = OUTPUT_DIR / "tradingqna_people.csv"
PEOPLE_XLSX = OUTPUT_DIR / "tradingqna_people.xlsx"

REQUEST_DELAY_SECONDS = 1.5
REQUEST_TIMEOUT_SECONDS = 45
MAX_RETRY_AFTER_SECONDS = 300
RETRIES = 2
SAVE_EVERY = 25
CHECKPOINT_EXPORT_EVERY = 100

# Keep this focused on market/trading categories for initial discovery.
TARGET_CATEGORY_NAMES = {
    # Original trading categories
    "Trading",
    "F&O",
    "Algos, strategies, code",
    "Nifty & Bank Nifty discussions",
    "Technical Analysis",
    "Stocks",
    "Fundamental Analysis",
    "Commodities",
    "Creators",
    "AMA (Ask Me Anything)",
    # Additional categories
    "General",
    "Streak - backtesting",
    "Personal finance",
    "Taxation",
    "Zerodha",
    "IPOs",
    "Industry Insights",
    "Sentinel",
    "Bitcoin & Crypto",
    "In The Money",
    "Aftermarket report",
    "Bonds",
    "Brokers",
    "Corporate actions",
}

TRADING_KEYWORDS = (
    "trader",
    "trading",
    "intraday",
    "swing",
    "scalp",
    "scalping",
    "options",
    "option",
    "futures",
    "f&o",
    "nifty",
    "bank nifty",
    "equity",
    "stocks",
    "commodity",
    "forex",
    "technical analysis",
    "fundamental analysis",
    "algo",
    "algorithmic",
    "price action",
    "derivatives",
    "option chain",
    "iv",
    "theta",
    "gamma",
    "hedging",
)

EMAIL_RE = re.compile(
    r"(?<![\w.+-])[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}(?![\w.-])"
)

# Used only on the public profile page's own bio text, not arbitrary posts.
PHONE_RE = re.compile(
    r"(?<!\d)(?:\+?91[\s.\-]?)?(?:0[\s.\-]?)?"
    r"(?:\(?\d{2,5}\)?[\s.\-]?)?"
    r"\d{3,5}[\s.\-]?\d{4,6}(?!\d)"
)

LABELED_PHONE_RE = re.compile(
    r"(?i)\b(?:phone|mobile|telephone|contact|call|whatsapp)\b"
    r"\s*(?:number|no\.?)?\s*[:\-]?\s*"
    r"(" + PHONE_RE.pattern + r")"
)

# A bio link is only treated as the person's own website when the same bio
# block also uses one of these labels -- an unlabeled external link (a cited
# article, a referral link, etc.) is not "clearly presented as the user's own
# public website" per policy.
WEBSITE_LABEL_RE = re.compile(
    r"(?i)\b(website|home\s*page|my\s+(?:site|blog|page)|portfolio)\b"
)

SOCIAL_DOMAINS = {
    "linkedin": ("linkedin.com",),
    "twitter": ("twitter.com", "x.com"),
    "instagram": ("instagram.com",),
    "youtube": ("youtube.com", "youtu.be"),
    "telegram": ("t.me", "telegram.me", "telegram.dog"),
    "whatsapp": ("wa.me", "whatsapp.com"),
}

# city/state/country are kept for schema stability but are always empty for
# this source: TradingQnA's crawler-safe profile HTML has no location field.
PEOPLE_COLUMNS = [
    "person_id",
    "name",
    "username",
    "email",
    "phone",
    "whatsapp",
    "telegram",
    "city",
    "state",
    "country",
    "bio",
    "website",
    "linkedin",
    "twitter",
    "instagram",
    "youtube",
    "trading_categories",
    "trading_keywords",
    "profile_url",
    "source",
    "topics_seen",
    "posts_seen",
]

DISCOVERY_COLUMNS = [
    "post_id",
    "person_id",
    "username",
    "name",
    "topic_id",
    "topic_title",
    "topic_url",
    "category",
    "post_number",
    "post_date",
    "post_url",
    "source",
]


# --------------------------------------------------------------------------- basic helpers

def banner(title: str) -> None:
    print(f"\n{'=' * 80}\n{title}\n{'=' * 80}")


def clean(value: object) -> str:
    if value is None:
        return ""
    text = html_lib.unescape(str(value))
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def stable_id(value: str) -> str:
    digest = hashlib.sha1(value.lower().encode("utf-8")).hexdigest()[:12]
    return f"TQNA-{digest}"


def absolute_url(href: str, base: str = BASE_URL) -> str:
    return urljoin(base, href.strip())


def same_host(url: str) -> bool:
    return urlsplit(url).netloc.lower() == urlsplit(BASE_URL).netloc.lower()


def normalized_url(url: str) -> str:
    parts = urlsplit(url)
    return urlunsplit(
        (
            parts.scheme,
            parts.netloc,
            parts.path.rstrip("/") or "/",
            parts.query,
            "",
        )
    )


def add_page_param(url: str, page: int) -> str:
    if page <= 0:
        return normalized_url(url)

    parts = urlsplit(url)
    params = dict(parse_qsl(parts.query, keep_blank_values=True))
    params["page"] = str(page)

    return urlunsplit(
        (
            parts.scheme,
            parts.netloc,
            parts.path,
            urlencode(params),
            "",
        )
    )


def extract_text(node: HtmlElement | None) -> str:
    if node is None:
        return ""
    try:
        return clean(" ".join(node.itertext()))
    except (AttributeError, TypeError):
        return clean(node.text_content() if hasattr(node, "text_content") else "")


def first_text(root: HtmlElement, xpaths: list[str]) -> str:
    for xpath in xpaths:
        try:
            nodes = root.xpath(xpath)
        except Exception:
            continue
        for node in nodes:
            value = extract_text(node) if hasattr(node, "itertext") else clean(node)
            if value:
                return value
    return ""


def uniq(values: list[str]) -> list[str]:
    seen: list[str] = []
    for value in values:
        value = clean(value)
        if value and value not in seen:
            seen.append(value)
    return seen


def uniq_join(values: list[str]) -> str:
    return "; ".join(uniq(values))


def keyword_list(texts: list[str]) -> list[str]:
    combined = " ".join(clean(t).lower() for t in texts if t)
    found = []
    for keyword in TRADING_KEYWORDS:
        if re.search(rf"(?<![a-z0-9]){re.escape(keyword)}(?![a-z0-9])", combined):
            found.append(keyword)
    return found


# --------------------------------------------------------------------------- cache

def empty_cache() -> dict:
    return {
        "categories": [],
        "category_pages": {},
        "topics": {},
        "profiles": {},
        "discovery": {},
    }


def load_cache(fresh: bool) -> dict:
    if fresh or not CACHE_PATH.exists():
        return empty_cache()

    try:
        data = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return empty_cache()

    fresh_shape = empty_cache()
    for key, default in fresh_shape.items():
        data.setdefault(key, default)
    return data


def save_cache(cache: dict) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    tmp = CACHE_PATH.with_suffix(".tmp")
    tmp.write_text(
        json.dumps(cache, ensure_ascii=False, separators=(",", ":")),
        encoding="utf-8",
    )
    tmp.replace(CACHE_PATH)


# --------------------------------------------------------------------------- robots

class RobotsGuard:
    def __init__(self) -> None:
        self.rp = robotparser.RobotFileParser()
        self.url = f"{BASE_URL}/robots.txt"

        headers = {
            "User-Agent": "TradingQnA-Public-Research/1.0",
            "Accept": "text/plain,*/*;q=0.8",
        }

        try:
            response = requests.get(self.url, headers=headers, timeout=30)
        except requests.RequestException as exc:
            raise PermissionError(f"Could not fetch robots.txt: {exc}. Stopping.") from exc

        if response.status_code != 200:
            raise PermissionError(
                f"robots.txt returned HTTP {response.status_code}. Stopping."
            )

        self.rp.set_url(self.url)
        self.rp.parse(response.text.splitlines())
        self.rp.modified()

        self.crawl_delay = self.rp.crawl_delay("*") or 0.0
        self.delay = max(REQUEST_DELAY_SECONDS, float(self.crawl_delay))

    def allowed(self, url: str) -> bool:
        return bool(self.rp.can_fetch("*", url))


def check_start_paths(robots: RobotsGuard) -> None:
    """
    Only check public HTML paths we actually intend to crawl.
    JSON/API endpoints are intentionally absent. Paths ending in "/" are
    representative path prefixes, not real IDs -- robots rules are path-based.
    """
    paths = ["/categories", "/latest", "/c/", "/t/", "/u/"]
    blocked = [path for path in paths if not robots.allowed(f"{BASE_URL}{path}")]

    if blocked:
        raise PermissionError(f"robots.txt disallows: {blocked}. Stopping.")


# --------------------------------------------------------------------------- HTTP client

class PublicHTMLClient:
    def __init__(self, robots: RobotsGuard) -> None:
        self.robots = robots
        self.session = requests.Session()
        self.requests_made = 0

        # Transparent identification; no browser impersonation or stealth headers.
        self.session.headers.update(
            {
                "User-Agent": "TradingQnA-Public-Research/1.0",
                "Accept": "text/html,application/xhtml+xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-IN,en;q=0.9",
            }
        )

    def get(self, url: str) -> requests.Response:
        url = normalized_url(url)

        if not same_host(url):
            raise PermissionError(f"Refusing off-site request: {url}")

        if not self.robots.allowed(url):
            raise PermissionError(f"robots.txt disallows: {url}")

        last_error: Exception | None = None

        for attempt in range(RETRIES + 1):
            time.sleep(self.robots.delay if attempt == 0 else self.robots.delay * 2)

            try:
                response = self.session.get(
                    url,
                    timeout=REQUEST_TIMEOUT_SECONDS,
                    allow_redirects=True,
                )
                self.requests_made += 1
            except requests.RequestException as exc:
                last_error = exc
                if attempt >= RETRIES:
                    break
                continue

            if not same_host(response.url):
                raise PermissionError(f"Request redirected off-site ({response.url}); stopping.")

            if response.status_code in (401, 403):
                raise PermissionError(
                    f"HTTP {response.status_code} from {url}. Access denied; stopping."
                )

            if response.status_code == 429:
                try:
                    wait = int(response.headers.get("Retry-After", "60") or 60)
                except ValueError:
                    wait = 60

                wait = min(max(wait, 1), MAX_RETRY_AFTER_SECONDS)

                if attempt < RETRIES:
                    print(f"  429 rate limit: waiting {wait}s as server requested")
                    time.sleep(wait)
                    continue

                raise PermissionError(
                    f"HTTP 429 persists for {url}; stopping without bypassing the limit."
                )

            try:
                response.raise_for_status()
            except requests.RequestException as exc:
                last_error = exc
                if attempt >= RETRIES:
                    break
                continue

            content_type = response.headers.get("Content-Type", "").lower()
            if "text/html" not in content_type and content_type:
                raise RuntimeError(
                    f"Unexpected content type {content_type!r} for {url}; expected HTML."
                )

            return response

        raise RuntimeError(f"Request failed: {url} -> {last_error}")


# --------------------------------------------------------------------------- category discovery

def discover_categories(client: PublicHTMLClient, cache: dict) -> list[dict]:
    banner("[1] PUBLIC CATEGORY HTML")

    if cache["categories"]:
        print(f"Using cached categories: {len(cache['categories'])}")
        return cache["categories"]

    response = client.get(CATEGORIES_URL)

    try:
        doc = lxml_html.fromstring(response.content)
    except (ValueError, TypeError) as exc:
        raise RuntimeError(f"Could not parse {CATEGORIES_URL}") from exc

    categories: list[dict] = []
    seen: set[str] = set()

    for anchor in doc.xpath("//a[@href]"):
        href = clean(anchor.get("href"))
        if not href:
            continue

        url = absolute_url(href)
        if not same_host(url):
            continue

        parts = urlsplit(url)
        match = re.fullmatch(r"/c/([^/]+)/(\d+)", parts.path)
        if not match:
            continue

        name = clean(anchor.text_content())
        if not name or name not in TARGET_CATEGORY_NAMES:
            continue

        key = parts.path
        if key in seen:
            continue

        seen.add(key)
        categories.append(
            {
                "name": name,
                "slug": match.group(1),
                "id": int(match.group(2)),
                "url": f"{BASE_URL}{parts.path}",
            }
        )

    categories.sort(key=lambda x: x["name"].lower())
    cache["categories"] = categories
    save_cache(cache)

    print(f"Selected categories: {len(categories)}")
    for category in categories:
        print(f"  {category['name']} -> {category['url']}")

    return categories


# --------------------------------------------------------------------------- topic-list HTML parsing

def topic_links_from_page(doc: HtmlElement, base_url: str) -> list[dict]:
    """
    Parse topic links from the public category HTML.

    We prefer the standard Discourse topic list rows and fall back to main-content
    topic links when the markup differs.
    """
    rows = doc.xpath(
        "//tr[contains(concat(' ', normalize-space(@class), ' '), ' topic-list-item ')]"
    )

    anchors = []
    if rows:
        for row in rows:
            anchors.extend(row.xpath(".//a[@href]"))
    else:
        anchors = doc.xpath(
            "//main//a[@href and contains(@href, '/t/')]"
            " | //*[@id='main-outlet']//a[@href and contains(@href, '/t/')]"
        )

    output: list[dict] = []
    seen: set[str] = set()

    for anchor in anchors:
        href = clean(anchor.get("href"))
        if not href:
            continue

        url = normalized_url(absolute_url(href, base_url))
        if not same_host(url):
            continue

        parts = urlsplit(url)
        if not parts.path.startswith("/t/"):
            continue

        # Ignore reply/post anchors and non-topic paths.
        path_parts = [p for p in parts.path.split("/") if p]
        if len(path_parts) < 2:
            continue

        # Discourse topic URLs are /t/<slug>/<id> or /t/<id>.
        numeric_id = next((p for p in reversed(path_parts) if p.isdigit()), "")
        if not numeric_id:
            continue

        title = clean(anchor.text_content()) or clean(anchor.get("title"))
        if not title:
            continue

        key = url.split("#", 1)[0]
        if key in seen:
            continue
        seen.add(key)

        # Try to grab reply/view counts from the same topic-list row.
        reply_count = ""
        view_count = ""
        parent_rows = anchor.xpath(
            "ancestor::tr[contains(concat(' ', normalize-space(@class), ' '), ' topic-list-item ')][1]"
        )
        if parent_rows:
            row_text = extract_text(parent_rows[0])
            # Keep these optional; not every layout exposes predictable selectors.
            nums = re.findall(r"\b\d[\d,]*\b", row_text)
            if len(nums) >= 2:
                reply_count = nums[0]
                view_count = nums[1]

        output.append(
            {
                "id": int(numeric_id),
                "title": title,
                "url": key,
                "reply_count": reply_count,
                "views": view_count,
            }
        )

    return output


def next_page_from_html(doc: HtmlElement, current_url: str) -> str | None:
    candidates: list[str] = []

    for node in doc.xpath(
        "//link[translate(@rel,'NEXT','next')='next']"
        " | //a[translate(@rel,'NEXT','next')='next']"
        " | //a[contains(@class,'next') and @href]"
    ):
        href = clean(node.get("href"))
        if href:
            candidates.append(absolute_url(href, current_url))

    # Fallback: pagination links whose visible label is "Next".
    for node in doc.xpath("//a[@href]"):
        label = clean(node.text_content()).lower()
        if label in {"next", "next page", "older"}:
            href = clean(node.get("href"))
            if href:
                candidates.append(absolute_url(href, current_url))

    for candidate in candidates:
        candidate = normalized_url(candidate)
        if same_host(candidate) and candidate != normalized_url(current_url):
            return candidate

    return None


def collect_topics(
    client: PublicHTMLClient,
    cache: dict,
    categories: list[dict],
    max_pages_per_category: int,
    max_topics: int | None,
) -> tuple[list[dict], int]:
    banner("[2] PUBLIC TOPIC LIST HTML")

    topics: dict[str, dict] = cache["topics"]
    pages_visited = 0

    for category in categories:
        if max_topics and len(topics) >= max_topics:
            break

        base_url = category["url"]
        print(f"\nCategory: {category['name']}")

        page = 0
        while page < max_pages_per_category:
            if max_topics and len(topics) >= max_topics:
                break

            page_url = add_page_param(base_url, page)
            page_key = normalized_url(page_url)

            # If we already visited this exact page, still keep scanning later pages
            # when the user increases max-pages, but don't request it twice.
            if page_key in cache["category_pages"]:
                cached_page = cache["category_pages"][page_key]
                for topic in cached_page.get("topics", []):
                    topics.setdefault(str(topic["id"]), {**topic, "category": category["name"]})
                next_url = cached_page.get("next_url")
                pages_visited += 1
                print(f"  Page {page}: cached | topics total: {len(topics):,}")
                if not next_url:
                    break
                page += 1
                continue

            try:
                response = client.get(page_url)
            except PermissionError:
                raise
            except RuntimeError as exc:
                print(f"  Page {page}: failed ({exc})")
                break

            try:
                doc = lxml_html.fromstring(response.content)
            except (ValueError, TypeError) as exc:
                print(f"  Page {page}: HTML parse failed ({exc})")
                break

            rows = topic_links_from_page(doc, page_url)
            next_url = next_page_from_html(doc, page_url)

            cached_topics = []
            for topic in rows:
                record = {
                    **topic,
                    "category": category["name"],
                    "created_at": "",
                    "last_posted_at": "",
                }
                topics.setdefault(str(topic["id"]), record)
                cached_topics.append(record)

                if max_topics and len(topics) >= max_topics:
                    break

            cache["category_pages"][page_key] = {
                "category": category["name"],
                "page": page,
                "topics": cached_topics,
                "next_url": next_url or "",
            }

            save_cache(cache)
            pages_visited += 1

            print(f"  Page {page}: {len(rows)} topics | unique total: {len(topics):,}")

            if max_topics and len(topics) >= max_topics:
                break

            if not next_url:
                break

            page += 1

    all_topics = list(topics.values())
    all_topics.sort(key=lambda x: (str(x.get("category", "")), int(x.get("id", 0))))

    if max_topics:
        all_topics = all_topics[:max_topics]

    print(f"\nTopic pages visited: {pages_visited:,}")
    print(f"Topics available in cache: {len(topics):,}")
    print(f"Topics selected for processing: {len(all_topics):,}")

    return all_topics, pages_visited


# --------------------------------------------------------------------------- topic detail HTML

def extract_topic_title(doc: HtmlElement, fallback: str) -> str:
    title = first_text(
        doc,
        [
            "//div[@id='topic-title']//h1",
            "//h1[contains(@class,'fancy-title')]",
            "//main//h1",
            "//h1",
        ],
    )
    return title or fallback


def topic_post_records(doc: HtmlElement) -> list[dict]:
    """
    Extract one record per real post on this topic HTML page.

    TradingQnA's crawler-safe topic HTML renders each post as
    `<div id="post_<n>" class="topic-body crawler-post">` with the author link
    inside `.crawler-post-meta .creator`. When that structure isn't present
    (an edge-case page or a template change), we fall back to scanning for
    public `/u/<username>` links so the topic still contributes discovery
    data -- but those fallback entries carry no verified post_id/post_number,
    so callers must not count them as confirmed "posts seen".
    """
    records: list[dict] = []
    seen_users: set[str] = set()

    posts = doc.xpath(
        "//div[contains(concat(' ', normalize-space(@class), ' '), ' crawler-post ')]"
    )

    for post in posts:
        author_links = post.xpath(
            ".//*[contains(concat(' ', normalize-space(@class), ' '), ' creator ')]"
            "//a[@href][1]"
        )
        if not author_links:
            continue

        href = clean(author_links[0].get("href"))
        match = re.search(r"/u/([^/?#]+)", urlsplit(absolute_url(href)).path)
        if not match:
            continue

        username = match.group(1)
        key = username.lower()
        if key in seen_users:
            continue
        seen_users.add(key)

        post_id = clean(post.get("id"))  # e.g. "post_3"
        post_number = post_id.split("_")[-1] if post_id.startswith("post_") else ""

        date_nodes = post.xpath(".//time[@datetime][1]")
        post_date = clean(date_nodes[0].get("datetime")) if date_nodes else ""

        records.append(
            {
                "post_id": post_id,
                "username": username,
                "name": extract_text(author_links[0]) or username,
                "post_number": post_number,
                "post_date": post_date,
            }
        )

    if records:
        return records

    # Fallback: scan for public /u/<username> links (e.g. @mentions, a changed
    # template). No post_id/post_number is available here.
    for link in doc.xpath("//a[@href]"):
        href = clean(link.get("href"))
        full_url = absolute_url(href)

        if not same_host(full_url):
            continue

        path = urlsplit(full_url).path
        match = re.fullmatch(r"/u/([^/?#]+)", path)
        if not match:
            continue

        username = match.group(1)
        key = username.lower()
        if key in seen_users:
            continue

        seen_users.add(key)
        records.append(
            {
                "post_id": "",
                "username": username,
                "name": clean(link.text_content()) or username,
                "post_number": "",
                "post_date": "",
            }
        )

    return records


def collect_topic_people(
    client: PublicHTMLClient,
    cache: dict,
    topics: list[dict],
    topic_pages: int,
) -> tuple[dict[str, dict], dict[str, dict]]:
    banner("[3] DISCOVER PUBLIC AUTHORS FROM TOPIC HTML")

    people: dict[str, dict] = {}
    discovery: dict[str, dict] = cache["discovery"]

    for idx, topic in enumerate(topics, 1):
        topic_id = int(topic["id"])
        topic_url = normalized_url(topic["url"])

        page = 0
        while page < max(topic_pages, 1):
            page_url = add_page_param(topic_url, page)

            cached = cache["topics"].get(str(topic_id), {})
            detail_pages = cached.setdefault("detail_pages", {})

            if str(page) in detail_pages:
                page_data = detail_pages[str(page)]
            else:
                try:
                    response = client.get(page_url)
                except PermissionError:
                    raise
                except RuntimeError as exc:
                    print(f"  Topic {topic_id}, page {page}: failed ({exc})")
                    break

                try:
                    doc = lxml_html.fromstring(response.content)
                except (ValueError, TypeError) as exc:
                    print(f"  Topic {topic_id}, page {page}: parse failed ({exc})")
                    break

                title = extract_topic_title(doc, topic.get("title", ""))
                posts = topic_post_records(doc)

                page_data = {
                    "title": title,
                    "posts": posts,
                    "url": normalized_url(page_url),
                    "next_page": bool(next_page_from_html(doc, page_url)),
                }

                detail_pages[str(page)] = page_data
                cache["topics"][str(topic_id)] = {
                    **cached,
                    **topic,
                    "category": topic.get("category", ""),
                    "detail_pages": detail_pages,
                }

                if page == 0:
                    cache["topics"][str(topic_id)]["title"] = title

                if idx % SAVE_EVERY == 0 or page > 0:
                    save_cache(cache)

            for post in page_data.get("posts", []):
                username = clean(post.get("username"))
                if not username:
                    continue

                person_key = username.lower()
                person = people.setdefault(
                    person_key,
                    {
                        "person_id": stable_id(username),
                        "name": clean(post.get("name")) or username,
                        "username": username,
                        "email": "",
                        "phone": "",
                        "whatsapp": "",
                        "telegram": "",
                        "city": "",
                        "state": "",
                        "country": "",
                        "bio": "",
                        "website": "",
                        "linkedin": "",
                        "twitter": "",
                        "instagram": "",
                        "youtube": "",
                        "trading_categories": [],
                        "trading_keywords": [],
                        "profile_url": f"{BASE_URL}/u/{username}",
                        "source": "TradingQnA",
                        "topics_seen": 0,
                        "posts_seen": 0,
                        "_topic_ids": set(),
                    },
                )

                if topic.get("category") and topic["category"] not in person["trading_categories"]:
                    person["trading_categories"].append(topic["category"])

                for keyword in keyword_list([topic.get("title", ""), topic.get("category", "")]):
                    if keyword not in person["trading_keywords"]:
                        person["trading_keywords"].append(keyword)

                person["_topic_ids"].add(topic_id)

                post_id = post.get("post_id")
                if post_id:
                    person["posts_seen"] += 1

                    dkey = str(post_id)
                    discovery[dkey] = {
                        "post_id": post_id,
                        "person_id": person["person_id"],
                        "username": username,
                        "name": person["name"],
                        "topic_id": topic_id,
                        "topic_title": topic.get("title", ""),
                        "topic_url": topic_url,
                        "category": topic.get("category", ""),
                        "post_number": post.get("post_number", ""),
                        "post_date": post.get("post_date", ""),
                        "post_url": (
                            f"{topic_url}/{post['post_number']}"
                            if post.get("post_number")
                            else topic_url
                        ),
                        "source": "TradingQnA",
                    }

            next_page = bool(page_data.get("next_page"))
            if not next_page:
                # HTML-only first page is enough for discovery unless caller explicitly asks
                # for additional topic pages.
                break

            page += 1

        if idx % SAVE_EVERY == 0 or idx == len(topics):
            for p in people.values():
                p["topics_seen"] = len(p["_topic_ids"])
            save_cache(cache)

            if idx % CHECKPOINT_EXPORT_EVERY == 0 or idx == len(topics):
                export_outputs(people, discovery)

            print(
                f"  Topics processed: {idx:,}/{len(topics):,} | "
                f"unique public authors: {len(people):,}"
            )

    for p in people.values():
        p["topics_seen"] = len(p["_topic_ids"])
        del p["_topic_ids"]

    return people, discovery


# --------------------------------------------------------------------------- profile HTML parsing

def public_profile_root(doc: HtmlElement) -> HtmlElement:
    nodes = doc.xpath(
        "//*[@id='main-outlet']"
        " | //main"
        " | //div[contains(concat(' ', normalize-space(@class), ' '), ' user-profile ')]"
    )
    return nodes[0] if nodes else doc


def profile_bio_nodes(root: HtmlElement) -> list[HtmlElement]:
    """
    Return the block-level elements that make up the user's own bio content.

    TradingQnA's crawler-safe profile HTML renders an avatar+username block
    (`div.user-crawler`) followed directly by whatever the user wrote as their
    bio (plain paragraphs, lists, etc.). Everything before that avatar block
    is page chrome (nav, banners), not user-published content.
    """
    avatar_block = root.xpath(
        ".//div[contains(concat(' ', normalize-space(@class), ' '), ' user-crawler ')]"
    )
    if avatar_block:
        return avatar_block[0].xpath("following-sibling::*")
    return list(root)


def extract_profile_links(bio_nodes: list[HtmlElement]) -> dict[str, str]:
    """
    Pull social/website links out of the user's own bio content only.

    A link is stored under a social field when its domain matches a known
    platform. A link is stored as "website" only when the same bio block also
    contains an explicit label (e.g. "Website:", "my blog") -- an unlabeled
    external link (a cited article, a referral link, etc.) is not clearly
    presented as the user's own site and is left out.
    """
    result: dict[str, str] = {}

    for node in bio_nodes:
        has_website_label = bool(WEBSITE_LABEL_RE.search(extract_text(node)))

        for anchor in node.xpath(".//a[@href]"):
            href = clean(anchor.get("href"))
            if not href or href.lower().startswith(("mailto:", "tel:", "javascript:", "#")):
                continue

            full = absolute_url(href, BASE_URL)
            if same_host(full):
                continue

            low = full.lower()
            matched_social = False
            for field, domains in SOCIAL_DOMAINS.items():
                if field not in result and any(domain in low for domain in domains):
                    result[field] = full
                    matched_social = True
                    break

            if not matched_social and has_website_label and "website" not in result:
                result["website"] = full

    return result


def parse_public_profile(html_bytes: bytes, username: str) -> dict:
    empty_profile = {
        "username": username,
        "name": username,
        "email": "",
        "phone": "",
        "whatsapp": "",
        "telegram": "",
        "city": "",
        "state": "",
        "country": "",
        "bio": "",
        "website": "",
        "linkedin": "",
        "twitter": "",
        "instagram": "",
        "youtube": "",
        "profile_url": f"{BASE_URL}/u/{username}",
    }

    try:
        doc = lxml_html.fromstring(html_bytes)
    except (ValueError, TypeError):
        return empty_profile

    root = public_profile_root(doc)
    bio_nodes = profile_bio_nodes(root)
    bio_text = clean(" ".join(extract_text(node) for node in bio_nodes))

    emails: list[str] = []
    phones: list[str] = []

    for node in bio_nodes:
        for anchor in node.xpath(".//a[starts-with(translate(@href,'MAILTO','mailto'),'mailto:')]"):
            emails.append(clean(anchor.get("href"))[7:].strip())

        for anchor in node.xpath(".//a[starts-with(translate(@href,'TEL','tel'),'tel:')]"):
            phones.append(clean(anchor.get("href"))[4:].strip())

    emails.extend(EMAIL_RE.findall(bio_text))

    # Prefer explicit public contact labels. Avoid treating dates, IDs, counts,
    # etc. as phone numbers.
    for match in LABELED_PHONE_RE.findall(bio_text):
        digits = re.sub(r"\D", "", match)
        if len(digits) >= 10:
            phones.append(match)

    links = extract_profile_links(bio_nodes)

    return {
        "username": username,
        "name": username,
        "email": uniq_join([e.lower() for e in emails]),
        "phone": uniq_join(phones),
        "whatsapp": links.get("whatsapp", ""),
        "telegram": links.get("telegram", ""),
        "city": "",
        "state": "",
        "country": "",
        "bio": bio_text,
        "website": links.get("website", ""),
        "linkedin": links.get("linkedin", ""),
        "twitter": links.get("twitter", ""),
        "instagram": links.get("instagram", ""),
        "youtube": links.get("youtube", ""),
        "profile_url": f"{BASE_URL}/u/{username}",
    }


def merge_profile_into_person(person: dict, profile: dict) -> None:
    for field in (
        "website",
        "linkedin",
        "twitter",
        "instagram",
        "youtube",
        "whatsapp",
        "telegram",
        "city",
        "state",
        "country",
        "bio",
    ):
        value = clean(profile.get(field))
        if value and not clean(person.get(field)):
            person[field] = value

    person["email"] = uniq_join([person.get("email", ""), profile.get("email", "")])
    person["phone"] = uniq_join([person.get("phone", ""), profile.get("phone", "")])


def enrich_public_profiles(
    client: PublicHTMLClient,
    robots: RobotsGuard,
    cache: dict,
    people: dict[str, dict],
    max_profiles: int | None,
) -> None:
    banner("[4] PUBLIC PROFILE HTML")

    usernames = sorted(people.keys())
    if max_profiles:
        usernames = usernames[:max_profiles]

    fetched = 0
    skipped_robots = 0

    for idx, person_key in enumerate(usernames, 1):
        person = people[person_key]
        username = person["username"]
        url = f"{BASE_URL}/u/{username}"

        if username.lower() in cache["profiles"]:
            merge_profile_into_person(person, cache["profiles"][username.lower()])
            continue

        if not robots.allowed(url):
            skipped_robots += 1
            continue

        try:
            response = client.get(url)
        except PermissionError:
            raise
        except RuntimeError as exc:
            print(f"  Profile {username}: failed ({exc})")
            continue

        profile = parse_public_profile(response.content, username)
        cache["profiles"][username.lower()] = profile
        merge_profile_into_person(person, profile)
        fetched += 1

        if idx % SAVE_EVERY == 0 or idx == len(usernames):
            save_cache(cache)

            if idx % CHECKPOINT_EXPORT_EVERY == 0 or idx == len(usernames):
                export_outputs(people, cache["discovery"])

            print(f"  Profiles fetched: {idx:,}/{len(usernames):,}")

    save_cache(cache)

    print(f"Profiles fetched this run: {fetched:,}")
    print(f"Profiles blocked by robots: {skipped_robots:,}")


# --------------------------------------------------------------------------- export

def people_dataframe(people: dict[str, dict]) -> pd.DataFrame:
    rows = []

    for person in people.values():
        row = {column: person.get(column, "") for column in PEOPLE_COLUMNS}
        row["trading_categories"] = uniq_join(person.get("trading_categories", []))
        row["trading_keywords"] = uniq_join(person.get("trading_keywords", []))
        rows.append(row)

    df = pd.DataFrame(rows, columns=PEOPLE_COLUMNS)

    if not df.empty:
        df = df.drop_duplicates(subset=["username"], keep="first")

    return df


def discovery_dataframe(discovery: dict[str, dict]) -> pd.DataFrame:
    return pd.DataFrame(list(discovery.values()), columns=DISCOVERY_COLUMNS)


def export_outputs(people: dict[str, dict], discovery: dict[str, dict]) -> None:
    banner("[5] EXPORT")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    people_df = people_dataframe(people)
    discovery_df = discovery_dataframe(discovery)

    try:
        people_df.to_csv(PEOPLE_CSV, index=False, encoding="utf-8-sig")

        with pd.ExcelWriter(PEOPLE_XLSX, engine="openpyxl") as writer:
            people_df.to_excel(writer, sheet_name="people", index=False)
            discovery_df.to_excel(writer, sheet_name="discovery", index=False)

            for ws in writer.sheets.values():
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions

    except PermissionError as exc:
        raise SystemExit(f"{exc}\nClose the CSV/Excel file and run again.") from exc

    print(f"People CSV : {PEOPLE_CSV}")
    print(f"People XLSX: {PEOPLE_XLSX}")
    print(f"People rows: {len(people_df):,}")
    print(f"Discovery rows: {len(discovery_df):,}")


# --------------------------------------------------------------------------- report

def report(
    client: PublicHTMLClient,
    people_df: pd.DataFrame,
    discovery_df: pd.DataFrame,
    pages: int,
) -> None:
    banner("SUMMARY")

    if people_df.empty:
        email_count = phone_count = both_count = either_count = website_count = 0
    else:
        email = people_df["email"].fillna("").str.strip().ne("")
        phone = people_df["phone"].fillna("").str.strip().ne("")
        email_count = int(email.sum())
        phone_count = int(phone.sum())
        both_count = int((email & phone).sum())
        either_count = int((email | phone).sum())
        website_count = int(people_df["website"].fillna("").str.strip().ne("").sum())

    print(f"HTTP requests this run    : {client.requests_made:,}")
    print(f"Topic-list pages visited  : {pages:,}")
    print(f"Unique people             : {len(people_df):,}")
    print(f"Unique discovery posts    : {len(discovery_df):,}")
    print(f"Public emails             : {email_count:,}")
    print(f"Public phones             : {phone_count:,}")
    print(f"Public email OR phone     : {either_count:,}")
    print(f"Public email AND phone    : {both_count:,}")
    print(f"Public website            : {website_count:,}")


# --------------------------------------------------------------------------- main

EPILOG = """
examples:
  # Small test run (fast, safe to run anytime)
  python -m scrapers.tradingqna --max-pages 2 --max-topics 50 --max-profiles 50

  # Resume a previous run using the local cache
  python -m scrapers.tradingqna --max-pages 30 --max-topics 7500 --max-profiles 12000

  # Larger collection run, ignoring any existing cache
  python -m scrapers.tradingqna --fresh --max-pages 30 --max-topics 7500 --max-profiles 12000

Only public HTML pages are requested (categories, topic lists, topics, user
profiles). robots.txt is checked before any request; no login, CAPTCHA
bypass, or JSON/API access is attempted.
"""


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="python -m scrapers.tradingqna",
        description=(
            "TradingQnA public HTML person discovery collector. "
            "Walks public category -> topic -> profile HTML pages on "
            "tradingqna.com and exports contact/social details that are "
            "explicitly published on each person's public profile."
        ),
        epilog=EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--max-pages",
        type=int,
        default=2,
        help="Maximum public category pages per category. Default: 2.",
    )
    parser.add_argument(
        "--max-topics",
        type=int,
        default=50,
        help="Maximum topics to process. Default: 50. Use 0 for no cap.",
    )
    parser.add_argument(
        "--topic-pages",
        type=int,
        default=1,
        help="Maximum HTML pages per topic. Default: 1.",
    )
    parser.add_argument(
        "--max-profiles",
        type=int,
        default=50,
        help="Maximum public profiles to fetch. Default: 50. Use 0 for no cap.",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Ignore existing local cache and start over.",
    )
    return parser


def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

    args = build_arg_parser().parse_args()

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    max_pages = max(args.max_pages, 1)
    max_topics = args.max_topics if args.max_topics > 0 else None
    max_profiles = args.max_profiles if args.max_profiles > 0 else None

    cache = load_cache(args.fresh)

    try:
        robots = RobotsGuard()
        check_start_paths(robots)

        client = PublicHTMLClient(robots)

        categories = discover_categories(client, cache)
        if not categories:
            raise SystemExit("No matching public trading categories found in /categories.")

        topics, pages = collect_topics(
            client,
            cache,
            categories,
            max_pages_per_category=max_pages,
            max_topics=max_topics,
        )

        # Person discovery remains HTML-only. We do not touch JSON/API endpoints.
        people, discovery = collect_topic_people(
            client,
            cache,
            topics,
            topic_pages=max(args.topic_pages, 1),
        )

        enrich_public_profiles(client, robots, cache, people, max_profiles=max_profiles)

        # Persist person discovery back into cache in a slim, JSON-safe form.
        # Sets are already removed by collect_topic_people.
        cache["discovery"] = discovery
        save_cache(cache)

        people_df = people_dataframe(people)
        discovery_df = discovery_dataframe(discovery)

        export_outputs(people, discovery)
        report(client, people_df, discovery_df, pages)

    except PermissionError as exc:
        save_cache(cache)
        print(f"\nSTOPPED: {exc}")
        print("No CAPTCHA/login/anti-bot/robots bypass was attempted.")
        raise SystemExit(2) from exc

    except KeyboardInterrupt:
        save_cache(cache)
        print("\nInterrupted. Progress saved; re-run to resume.")
        raise SystemExit(130) from None

    except Exception as exc:
        save_cache(cache)
        print(f"\nERROR: {exc}")
        print("Progress was saved to the cache.")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
