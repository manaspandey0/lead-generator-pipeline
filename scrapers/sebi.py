"""
scrapers/sebi.py - SEBI Registered Stock Brokers (equity segment): ALL pages -> CSV + Excel.

    python -m scrapers.sebi

Pagination = SEBI's own XHR call made by searchFormFpi():
    POST /sebiweb/ajax/other/getintmfpiinfo.jsp
    nextValue, next      <- the arguments of each page's own "Next >>" link
    intmId=30, language=2, other search fields <- page-1 form values (blank if absent)

Every page must continue exactly where the previous one ended (range text
"26 to 50 of 4991 ..."), otherwise the run FAILS with the reason.
Export happens only when fetched == SEBI's reported total.
Ordinary HTTP requests only. No CAPTCHA solving, no anti-bot bypassing.
"""
from __future__ import annotations

import math
import re
import sys
import time
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup, Tag

SOURCE = "SEBI - Registered Stock Brokers (Equity Segment)"
URL = "https://www.sebi.gov.in/sebiweb/other/OtherAction.do?doRecognisedFpi=yes&intmId=30"
PAGER_ENDPOINT = "/sebiweb/ajax/other/getintmfpiinfo.jsp"
PARSER = "html.parser"

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_DIR = ROOT / "output"
CSV_PATH = OUTPUT_DIR / "sebi_leads.csv"
XLSX_PATH = OUTPUT_DIR / "sebi_leads.xlsx"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-IN,en;q=0.9",
}
REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
DELAY_SECONDS = 1.5  # polite gap between page requests

# Parameters sent by searchFormFpi(), in its order. Values not listed in
# FIXED_PARAMS are taken from the page-1 form (by name or id), else blank.
PAGER_PARAMS = [
    "nextValue", "next", "intmId", "contPer", "name", "regNo", "email", "location",
    "exchange", "affiliate", "alp", "language", "model", "esgCategory", "doDirect", "intmIds",
]
FIXED_PARAMS = {"intmId": "30", "language": "2"}

RANGE_RE = re.compile(r"(\d[\d,]*)\s*to\s*(\d[\d,]*)\s*of\s*(\d[\d,]*)\s*records", re.I)
CALL_RE = re.compile(r"searchFormFpi\s*\(([^)]*)\)")
ARG_RE = re.compile(r"'([^']*)'|\"([^\"]*)\"|([^,\s]+)")
CONTROL_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")  # Excel rejects these
# Normalised SEBI label -> field key. Unknown labels are kept (snake_cased), never dropped.
LABEL_KEYS = {
    "name": "name",
    "trade name": "trade_name",
    "registration no": "registration_no",
    "e-mail": "email",
    "email": "email",
    "telephone": "telephone",
    "address": "address",
    "validity": "validity",
    "exchange name": "exchange_name",
}
KNOWN_FIELDS = ["name", "trade_name", "registration_no", "email", "telephone",
                "address", "validity", "exchange_name"]


class ScrapeError(RuntimeError):
    """The full dataset could not be collected with ordinary HTTP requests."""


# --------------------------------------------------------------------------- http
def http(session: requests.Session, method: str, url: str, data: dict | None = None,
         ajax: bool = False) -> str:
    headers = dict(HEADERS)
    if ajax:
        headers.update({"X-Requested-With": "XMLHttpRequest", "Referer": URL})
    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = session.request(method, url, headers=headers, data=data, timeout=REQUEST_TIMEOUT)
            if resp.status_code in (401, 403, 429):
                raise ScrapeError(f"HTTP {resp.status_code} from {url}: blocked or rate-limited. Not bypassing.")
            resp.raise_for_status()
            if not resp.encoding or resp.encoding.lower() == "iso-8859-1":
                resp.encoding = resp.apparent_encoding
            return resp.text
        except requests.RequestException as exc:
            last_error = exc
            if attempt < MAX_RETRIES:
                time.sleep(DELAY_SECONDS * 2 * attempt)
    raise ScrapeError(f"Request to {url} failed after {MAX_RETRIES} attempts: {last_error}")


# --------------------------------------------------------------------------- field split
def label_key(label: str) -> str:
    norm = re.sub(r"\s+", " ", label.strip().lower())
    norm = re.sub(r"[\s:.]+$", "", norm)
    return LABEL_KEYS.get(norm, re.sub(r"[^a-z0-9]+", "_", norm).strip("_"))


def split_field(card_view: Tag) -> tuple[str, str] | None:
    """One div.card-view -> (label, value). Value lives in the child with class 'value'."""
    value_el = card_view.find(class_="value")
    if value_el is not None:
        value_ids = {id(s) for s in value_el.strings}
        label = " ".join(
            s.strip() for s in card_view.strings if id(s) not in value_ids and s.strip()
        )
        value = value_el.get_text(" ", strip=True)
        if not label:  # label may be carried in an attribute instead of text
            label = card_view.get("data-label") or value_el.get("data-label") or value_el.get("title") or ""
    else:
        parts = list(card_view.stripped_strings)
        if not parts:
            return None
        label, value = parts[0], " ".join(parts[1:])
    return (label, value) if label else None


# --------------------------------------------------------------------------- parse
def parse_cards(html: str, source_url: str = URL, verbose: bool = False) -> list[dict]:
    """
    Every div.card-view is ONE field. A new record starts at each 'Name' label,
    so records with missing fields (fewer than 8 card-views) still parse correctly.
    Works on the full page and on the pager's HTML fragment.
    """
    soup = BeautifulSoup(html, PARSER)
    scope = soup.find(id="ajax_cat") or soup
    fields = [cv for cv in scope.select("div.card-view") if not cv.select_one("div.card-view")]

    records: list[dict] = []
    current: dict | None = None
    orphans = unsplittable = 0

    for cv in fields:
        pair = split_field(cv)
        if pair is None:
            unsplittable += 1
            continue
        key, value = label_key(pair[0]), pair[1]
        if key == "name":
            if current:
                records.append(current)
            current = {}
        if current is None:  # field seen before the first 'Name'
            orphans += 1
            continue
        current[key] = f"{current[key]} | {value}" if key in current and value else value
    if current:
        records.append(current)

    for r in records:
        r["source"] = SOURCE
        r["source_url"] = source_url

    if verbose:
        print(f"Fields scanned: {len(fields)} | unsplittable: {unsplittable} | before first Name: {orphans}")
    return records


def page_range(html: str) -> tuple[int, int, int] | None:
    m = RANGE_RE.search(BeautifulSoup(html, PARSER).get_text(" ", strip=True))
    return tuple(int(g.replace(",", "")) for g in m.groups()) if m else None


def verify(records: list[dict], html: str) -> bool:
    """Page-level check: parsed count must equal the page's own 'X to Y' range."""
    rng = page_range(html)
    expected = rng[1] - rng[0] + 1 if rng else None
    no_name = sum(1 for r in records if not r.get("name"))
    no_reg = sum(1 for r in records if not r.get("registration_no"))
    count_ok = expected is not None and len(records) == expected
    if not count_ok or no_name or no_reg:
        print(f"Page check failed: range={rng} parsed={len(records)} "
              f"missing_name={no_name} missing_reg={no_reg}")
    return count_ok and no_name == 0 and no_reg == 0


# --------------------------------------------------------------------------- pagination
def form_lookup(soup: BeautifulSoup) -> dict[str, str]:
    """Current values of the page's form fields, addressable by name and by id."""
    out: dict[str, str] = {}
    for el in soup.find_all(["input", "select", "textarea"]):
        typ = (el.get("type") or "").lower()
        if typ in ("submit", "button", "image", "reset", "file"):
            continue
        if typ in ("checkbox", "radio") and not el.has_attr("checked"):
            continue
        if el.name == "select":
            opt = el.find("option", selected=True) or el.find("option")
            value = opt.get("value", opt.get_text(strip=True)) if opt else ""
        elif el.name == "textarea":
            value = el.get_text()
        else:
            value = el.get("value", "")
        for key in (el.get("name"), el.get("id")):
            if key and key not in out:
                out[key] = value
    return out


def _call_args(inner: str) -> list[str]:
    return [a or b or c for a, b, c in ARG_RE.findall(inner)]


def _element_label(el: Tag) -> str:
    """Visible text, or title/alt/value/class when the link is an icon."""
    parts = [el.get_text(" ", strip=True), el.get("title", ""), el.get("aria-label", ""),
             el.get("value", ""), " ".join(el.get("class", []))]
    img = el.find("img")
    if img is not None:
        parts += [img.get("alt", ""), img.get("title", ""), img.get("src", "")]
    return " ".join(p for p in parts if p).lower()


def pager_calls(html: str) -> list[tuple[str, list[str]]]:
    """Every searchFormFpi(...) call with the label of the control that carries it."""
    soup = BeautifulSoup(html, PARSER)
    calls = []
    for el in soup.find_all(True):
        for attr in ("href", "onclick"):
            m = CALL_RE.search(str(el.get(attr, "")))
            if m:
                calls.append((_element_label(el), _call_args(m.group(1))))
    if not calls:  # pager written by a script: label = text right after the call
        for m in CALL_RE.finditer(html):
            after = re.sub(r"<[^>]+>", " ", html[m.end(): m.end() + 300]).split("searchFormFpi")[0]
            calls.append((re.sub(r"\s+", " ", after).strip().lower()[:40], _call_args(m.group(1))))
    return calls


def next_link_args(html: str) -> tuple[str, str] | None:
    """(next, nextValue) from the page's own 'Next' searchFormFpi(...) control."""
    for label, args in pager_calls(html):
        if len(args) >= 2 and ("next" in label or label.strip() in {">", ">>", "»", "›"}):
            return args[0], args[1]
    return None

def build_payload(form: dict[str, str], next_flag: str, direct: str, next_value: str) -> dict[str, str]:
    """searchFormFpi() parameters: Next link gives next + doDirect; nextValue is the page counter."""
    data = {name: FIXED_PARAMS.get(name, form.get(name, "")) for name in PAGER_PARAMS}
    data["next"] = next_flag
    data["doDirect"] = direct
    data["nextValue"] = next_value
    return data


# --------------------------------------------------------------------------- dedupe
def _norm(v) -> str:
    return re.sub(r"[^A-Z0-9]+", " ", str(v or "").upper()).strip()


def dedupe_key(r: dict) -> str:
    reg, ex = _norm(r.get("registration_no")), _norm(r.get("exchange_name"))
    if reg:
        return f"REG|{reg}|{ex}"
    return f"NAE|{_norm(r.get('name'))}|{_norm(r.get('address'))}|{ex}"


def deduplicate(records: list[dict]) -> list[dict]:
    unique: dict[str, dict] = {}
    for r in records:
        unique.setdefault(dedupe_key(r), r)
    return list(unique.values())


# --------------------------------------------------------------------------- scrape
def scrape_all() -> tuple[list[dict], int]:
    session = requests.Session()

    html = http(session, "GET", URL)
    rng = page_range(html)
    records = parse_cards(html)
    if rng is None or not verify(records, html):
        raise ScrapeError("Page 1 did not verify (range text or 25/25 parse). Parser/layout changed.")
    start, end, total = rng
    total_pages = math.ceil(total / (end - start + 1))
    print(f"Page 1: {len(records)} records  (rows {start}-{end} of {total:,})")

    fetched = list(records)
    seen = {dedupe_key(r) for r in records}
    if end >= total:
        return fetched, total

    form = form_lookup(BeautifulSoup(html, PARSER))
    link = next_link_args(html)
    if link is None:
        found = sorted({f"{label[:30]!r} -> searchFormFpi({', '.join(map(repr, args))})"
                        for label, args in pager_calls(html)})
        raise ScrapeError("Page 1 has no 'Next' searchFormFpi(...) control. Calls found on the page:\n  "
                          + ("\n  ".join(found[:15]) or "none"))
    next_flag, direct = link
    endpoint = urljoin(URL, PAGER_ENDPOINT)
    print(f"Pager: POST {endpoint}  (Next link: next={next_flag!r}, doDirect={direct!r})")

    # Page 2: find the nextValue SEBI expects. Each candidate is accepted only if
    # SEBI itself answers with rows starting at end+1 (e.g. "26 to 50 of 4991").
    candidates: list[str] = []
    for cand in (form.get("nextValue", ""), "1", "2", "0"):
        if re.fullmatch(r"\d+", cand or "") and cand not in candidates:
            candidates.append(cand)
    tried, first_value, page_html = [], None, ""
    for cand in candidates:
        time.sleep(DELAY_SECONDS)
        page_html = http(session, "POST", endpoint,
                         data=build_payload(form, next_flag, direct, cand), ajax=True)
        r2, n2 = page_range(page_html), len(parse_cards(page_html))
        tried.append(f"nextValue={cand}: {n2} records, range={r2}")
        if n2 and r2 and r2[0] == end + 1:
            first_value = int(cand)
            break
    if first_value is None:
        raise ScrapeError(f"SEBI returned no page starting at row {end + 1} for any nextValue:\n  "
                          + "\n  ".join(tried))
    print(f"nextValue for page 2 = {first_value} (verified by SEBI's range text)")

    page, expected_start = 2, end + 1
    while True:
        recs = parse_cards(page_html)
        rng = page_range(page_html)
        if not recs:
            print(f"Page {page}: 0 records - stopping")
            break
        if rng and rng[0] != expected_start:
            raise ScrapeError(
                f"Page {page} returned rows {rng[0]}-{rng[1]}, expected rows from {expected_start} "
                f"(sent nextValue={first_value + page - 2}). SEBI did not advance the listing."
            )
        if not any(dedupe_key(r) not in seen for r in recs):
            print(f"Page {page}: {len(recs)} records, none new - stopping")
            break
        print(f"Page {page}: {len(recs)} records")
        fetched.extend(recs)
        seen.update(dedupe_key(r) for r in recs)
        expected_start = (rng[1] if rng else expected_start + len(recs) - 1) + 1
        if expected_start > total:
            break

        page += 1
        if page > total_pages + 2:
            raise ScrapeError(f"Exceeded {total_pages + 2} pages without reaching {total:,} records.")
        time.sleep(DELAY_SECONDS)
        page_html = http(session, "POST", endpoint,
                         data=build_payload(form, next_flag, direct, str(first_value + page - 2)), ajax=True)

    return fetched, total


# --------------------------------------------------------------------------- export
def export(records: list[dict]) -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    extra = sorted({k for r in records for k in r} - set(KNOWN_FIELDS) - {"source", "source_url"})
    columns = KNOWN_FIELDS + extra + ["source", "source_url"]
    df = pd.DataFrame(records, columns=columns).fillna("")
    df = df.map(lambda v: CONTROL_CHARS_RE.sub(" ", v) if isinstance(v, str) else v)
    try:
        df.to_csv(CSV_PATH, index=False, encoding="utf-8-sig")
        with pd.ExcelWriter(XLSX_PATH, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="sebi_brokers", index=False)
            ws = writer.sheets["sebi_brokers"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for i, col in enumerate(df.columns, 1):
                width = max(len(col), int(df[col].astype(str).str.len().quantile(0.95)) if len(df) else 0)
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(width + 2, 60)
    except PermissionError as exc:
        raise ScrapeError(f"{exc}. Close the file in Excel and run again.") from exc


# --------------------------------------------------------------------------- main
def main() -> None:
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except AttributeError:
        pass

    t0 = time.time()
    try:
        fetched, expected = scrape_all()
    except ScrapeError as exc:
        print(f"\nFAIL: {exc}")
        print("Nothing exported. Only complete, verified runs are written to output/.")
        sys.exit(1)

    unique = deduplicate(fetched)
    missing_name = sum(1 for r in unique if not r.get("name"))
    missing_reg = sum(1 for r in unique if not r.get("registration_no"))
    complete = len(fetched) == expected

    print("\n" + "=" * 60)
    print(f"Total fetched  : {len(fetched):,}")
    print(f"Total unique   : {len(unique):,}")
    print(f"Total expected : {expected:,}")
    print("=" * 60)
    print(f"Fetched == expected     : {'OK' if complete else 'MISMATCH'}")
    print(f"Missing names           : {missing_name}")
    print(f"Missing registration no : {missing_reg}")
    print(f"Duplicates removed      : {len(fetched) - len(unique)}")
    print(f"Time                    : {time.time() - t0:.0f}s")

    if not complete:
        print(f"\nFAIL: fetched {len(fetched):,} of {expected:,} records. Nothing exported.")
        sys.exit(1)

    try:
        export(unique)
    except ScrapeError as exc:
        print(f"\nFAIL: {exc}")
        sys.exit(1)
    print(f"\nCSV   : {CSV_PATH}")
    print(f"Excel : {XLSX_PATH}")
    print("RESULT: PASS")


if __name__ == "__main__":
    main()