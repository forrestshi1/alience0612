"""Full pipeline for the Soumu radio-license sales-list project.

Covers all four client queries, paginated, deduplicated, then enriches every
company with all requested fields using the proven logic in
tele_soumu_oneclick_enrich.py (NTA / gBizINFO / IRBank / official site crawl /
PDF / web search), running enrichment concurrently with per-host politeness.

Usage:
    python tele_soumu_full_pipeline.py collect [--queries q1,q2,q3,q4] [--limit-pages N]
    python tele_soumu_full_pipeline.py enrich  [--workers 8] [--limit N] [--no-search-web]
    python tele_soumu_full_pipeline.py export
    python tele_soumu_full_pipeline.py all [--quick-test]

All state lives under ./pipeline_out/ and every phase resumes where it stopped:
    raw_records.csv     every license record scraped (audit trail)
    collect_state.json  per-query pagination cursor
    companies_master.csv deduplicated company list
    enriched_full.csv   enrichment output, appended row by row (resume key)
    deliverable.csv/.xlsx client-facing export with Japanese headers

Optional: set GBIZINFO_API_TOKEN to use the free gBizINFO REST API
(https://info.gbiz.go.jp/hojin/swagger-ui/index.html) which markedly improves
資本金 / 従業員数 / 決算月 fill rates.
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import threading
import time
import urllib.parse
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from bs4 import BeautifulSoup

sys.path.insert(0, str(Path(__file__).parent))
from tele_soumu_oneclick_enrich import (  # noqa: E402
    FINAL_FRONT_COLUMNS,
    HEADERS,
    SOUmu_SEARCH,
    classify_entity,
    clean,
    enrich_row,
    ensure_schema,
    format_xlsx,
    parse_name_and_corporate_number,
    set_field,
)

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

OUT_DIR = Path("pipeline_out")
RAW_CSV = OUT_DIR / "raw_records.csv"
STATE_JSON = OUT_DIR / "collect_state.json"
MASTER_CSV = OUT_DIR / "companies_master.csv"
ENRICHED_CSV = OUT_DIR / "enriched_full.csv"
DELIVER_PREFIX = OUT_DIR / "deliverable"

PAGE_SIZE = 100

PREFECTURES = {
    "08": "茨城県", "10": "群馬県", "11": "埼玉県", "12": "千葉県", "13": "東京都",
    "14": "神奈川県", "22": "静岡県", "23": "愛知県", "26": "京都府", "27": "大阪府",
    "29": "奈良県",
}
BUREAUS = {"A": "関東", "C": "東海", "E": "近畿"}


def query_units() -> list[dict]:
    units: list[dict] = []
    # ① 簡易無線局 × 11 prefectures
    for code, pref in PREFECTURES.items():
        units.append({
            "key": f"q1_{code}", "query": "q1", "region": pref,
            "params": {"SK": "2", "DC": str(PAGE_SIZE), "pageID": "3", "CONFIRM": "1",
                       "SelectID": "1", "SelectOW": "03", "HC": code},
            "select_id": "1",
        })
    # ② デジタル簡易無線局（登録局） × 3 bureaus
    for it, region in BUREAUS.items():
        units.append({
            "key": f"q2_{it}", "query": "q2", "region": region,
            "params": {"pageID": "3", "SelectID": "6", "CONFIRM": "1", "OW2": "008",
                       "IT": it, "DC": str(PAGE_SIZE)},
            "select_id": "6",
        })
    # ③ 陸上移動局（包括免許） 897.5MHz × 3 bureaus — must run before q4
    for it, region in BUREAUS.items():
        units.append({
            "key": f"q3_{it}", "query": "q3", "region": region,
            "params": {"pageID": "3", "SelectID": "5", "CONFIRM": "1", "OW": "ML 1",
                       "IT": it, "FF": "897.5", "TF": "897.5", "HZ": "3", "DC": str(PAGE_SIZE)},
            "select_id": "5",
        })
    # ④ 陸上移動局（包括免許） 全周波数 × 3 bureaus (③のDFCDを除外)
    for it, region in BUREAUS.items():
        units.append({
            "key": f"q4_{it}", "query": "q4", "region": region,
            "params": {"pageID": "3", "SelectID": "5", "CONFIRM": "1", "OW": "ML 1",
                       "IT": it, "HZ": "3", "DC": str(PAGE_SIZE)},
            "select_id": "5",
        })
    return units


class ThreadSafeFetcher:
    """requests-based fetcher: one session per thread, per-host rate limiting."""

    HOST_INTERVALS = {
        "www.tele.soumu.go.jp": 0.7,
        "www.houjin-bangou.nta.go.jp": 0.5,
        "info.gbiz.go.jp": 0.5,
        "irbank.net": 0.5,
        "duckduckgo.com": 3.0,
        "html.duckduckgo.com": 3.0,
        "www.bing.com": 3.0,
    }

    def __init__(self, default_interval: float = 0.35, timeout: int = 18, retries: int = 2) -> None:
        self.default_interval = default_interval
        self.timeout = timeout
        self.retries = retries
        self._local = threading.local()
        self._hosts: dict[str, dict] = {}
        self._hosts_guard = threading.Lock()

    # -- interface compatible with tele_soumu_oneclick_enrich.Fetcher --
    @property
    def sleep_seconds(self) -> float:
        return self.default_interval

    def _session(self) -> requests.Session:
        if not hasattr(self._local, "session"):
            s = requests.Session()
            s.headers.update(HEADERS)
            self._local.session = s
        return self._local.session

    def _throttle(self, url: str) -> None:
        host = urllib.parse.urlparse(url).netloc
        interval = self.HOST_INTERVALS.get(host, self.default_interval)
        with self._hosts_guard:
            entry = self._hosts.setdefault(host, {"lock": threading.Lock(), "ts": 0.0})
        with entry["lock"]:
            wait = interval - (time.time() - entry["ts"])
            if wait > 0:
                time.sleep(wait)
            entry["ts"] = time.time()

    def fetch(self, url: str, *, allow_insecure_retry: bool = True):
        from tele_soumu_oneclick_enrich import FetchResult, ensure_url

        url = ensure_url(url)
        if not url:
            return FetchResult(url, url, "", 0, "empty url")
        last_error = ""
        for attempt in range(self.retries + 1):
            self._throttle(url)
            try:
                response = self._session().get(url, timeout=self.timeout, allow_redirects=True, verify=True)
            except requests.exceptions.SSLError as exc:
                if not allow_insecure_retry:
                    last_error = str(exc)
                    continue
                try:
                    self._throttle(url)
                    response = self._session().get(url, timeout=self.timeout, allow_redirects=True, verify=False)
                except Exception as inner_exc:
                    last_error = str(inner_exc)
                    time.sleep(0.6 * (attempt + 1))
                    continue
            except Exception as exc:
                last_error = str(exc)
                time.sleep(0.6 * (attempt + 1))
                continue
            if response.status_code >= 500 and attempt < self.retries:
                time.sleep(0.8 * (attempt + 1))
                continue
            if response.status_code >= 400:
                return FetchResult(url, response.url, "", response.status_code, f"HTTP {response.status_code}")
            response.encoding = response.apparent_encoding or response.encoding
            return FetchResult(url, response.url, response.text, response.status_code, "")
        return FetchResult(url, url, "", 0, last_error)

    def fetch_bytes(self, url: str, *, max_bytes: int = 15_000_000):
        from tele_soumu_oneclick_enrich import ensure_url

        url = ensure_url(url)
        if not url:
            return b"", url, "empty url"
        last_error = ""
        for attempt in range(self.retries + 1):
            self._throttle(url)
            try:
                response = self._session().get(url, timeout=self.timeout, allow_redirects=True, verify=True)
            except requests.exceptions.SSLError:
                try:
                    self._throttle(url)
                    response = self._session().get(url, timeout=self.timeout, allow_redirects=True, verify=False)
                except Exception as exc:
                    last_error = str(exc)
                    time.sleep(0.6 * (attempt + 1))
                    continue
            except Exception as exc:
                last_error = str(exc)
                time.sleep(0.6 * (attempt + 1))
                continue
            if response.status_code >= 400:
                return b"", response.url, f"HTTP {response.status_code}"
            content = response.content
            if len(content) > max_bytes:
                return b"", response.url, f"too large: {len(content)} bytes"
            return content, response.url, ""
        return b"", url, last_error

    def fetch_json(self, url: str, extra_headers: dict[str, str]) -> dict:
        self._throttle(url)
        try:
            response = self._session().get(url, timeout=self.timeout, headers=extra_headers)
            if response.status_code >= 400:
                return {}
            return response.json()
        except Exception:
            return {}


# ---------------------------------------------------------------- collect ---

RAW_FIELDS = ["query", "unit", "region", "company_name", "corporate_number",
              "office_location", "purpose", "license_date", "dfcd", "raw_name"]


def build_url(params: dict[str, str], sc: int) -> str:
    merged = {**params, "SC": str(sc)}
    query = urllib.parse.urlencode(merged).replace("+", "%20")
    return f"{SOUmu_SEARCH}?{query}"


def parse_list_page(html: str, select_id: str) -> tuple[list[dict], int]:
    soup = BeautifulSoup(html or "", "html.parser")
    total = 0
    m = re.search(r"([\d,]+)\s*件中", html or "")
    if m:
        total = int(m.group(1).replace(",", ""))
    records: list[dict] = []
    for tr in soup.select("tr.m-table-sort__row"):
        tds = tr.find_all("td")
        if not tds:
            continue
        raw_name = clean(tds[0].get_text(" ", strip=True))
        name, number = parse_name_and_corporate_number(raw_name)
        link = tr.find("a", href=True)
        dfcd = ""
        if link:
            dm = re.search(r"DFCD=(\d+)", link["href"])
            dfcd = dm.group(1) if dm else ""
        cols = [clean(td.get_text(" ", strip=True)) for td in tds[1:4]]
        cols += [""] * (3 - len(cols))
        if select_id == "6":
            office_location, purpose, license_date = "", "デジタル簡易無線局(登録局)", cols[1]
        else:
            office_location, purpose, license_date = cols[0], cols[1], cols[2]
        records.append({
            "company_name": name,
            "corporate_number": number,
            "office_location": office_location,
            "purpose": purpose,
            "license_date": license_date,
            "dfcd": dfcd,
            "raw_name": raw_name,
        })
    return records, total


def is_masked(name: str) -> bool:
    return not name.strip("*＊ ")


def load_state() -> dict:
    if STATE_JSON.exists():
        return json.loads(STATE_JSON.read_text(encoding="utf-8"))
    return {}


def save_state(state: dict) -> None:
    STATE_JSON.write_text(json.dumps(state, ensure_ascii=False, indent=1), encoding="utf-8")


def q3_dfcd_set() -> set[str]:
    if not RAW_CSV.exists():
        return set()
    dfcds: set[str] = set()
    with RAW_CSV.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            if row.get("query") == "q3" and row.get("dfcd"):
                dfcds.add(row["dfcd"])
    return dfcds


def collect(queries: list[str], limit_pages: int, sleep: float) -> None:
    OUT_DIR.mkdir(exist_ok=True)
    fetcher = ThreadSafeFetcher(default_interval=sleep)
    state = load_state()
    units = [u for u in query_units() if u["query"] in queries]

    new_file = not RAW_CSV.exists()
    raw_f = RAW_CSV.open("a", encoding="utf-8-sig", newline="")
    writer = csv.DictWriter(raw_f, fieldnames=RAW_FIELDS, extrasaction="ignore")
    if new_file:
        writer.writeheader()

    exclude_dfcds: set[str] = set()
    try:
        for unit in units:
            key = unit["key"]
            unit_state = state.get(key, {"next_sc": 1, "done": False})
            if unit_state.get("done"):
                print(f"[{key}] already done, skip")
                continue
            if unit["query"] == "q4" and not exclude_dfcds:
                exclude_dfcds = q3_dfcd_set()
                print(f"[q4] excluding {len(exclude_dfcds)} q3 records (③以外)")
            sc = unit_state["next_sc"]
            pages_done = 0
            while True:
                result = fetcher.fetch(build_url(unit["params"], sc))
                records, total = parse_list_page(result.html, unit["select_id"])
                if not records:
                    if result.error:
                        print(f"[{key}] SC={sc} fetch error: {result.error}; retrying once")
                        time.sleep(3)
                        result = fetcher.fetch(build_url(unit["params"], sc))
                        records, total = parse_list_page(result.html, unit["select_id"])
                    if not records:
                        print(f"[{key}] SC={sc}: no rows (total={total}); marking done")
                        unit_state["done"] = True
                        break
                kept = 0
                for record in records:
                    if is_masked(record["company_name"]):
                        continue
                    if unit["query"] == "q4" and record["dfcd"] in exclude_dfcds:
                        continue
                    writer.writerow({"query": unit["query"], "unit": key, "region": unit["region"], **record})
                    kept += 1
                raw_f.flush()
                pages_done += 1
                print(f"[{key}] SC={sc} rows={len(records)} kept={kept} total={total}")
                sc += PAGE_SIZE
                unit_state["next_sc"] = sc
                state[key] = unit_state
                save_state(state)
                if sc > total or len(records) < PAGE_SIZE:
                    unit_state["done"] = True
                    break
                if limit_pages and pages_done >= limit_pages:
                    print(f"[{key}] page limit reached, will resume at SC={sc}")
                    break
            state[key] = unit_state
            save_state(state)
    finally:
        raw_f.close()
    build_master()


def build_master() -> None:
    if not RAW_CSV.exists():
        print("no raw records yet")
        return
    companies: dict[str, dict] = {}
    with RAW_CSV.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            name = clean(row.get("company_name", ""))
            if is_masked(name):
                continue
            key = row.get("corporate_number") or f"name:{name}"
            entry = companies.setdefault(key, {
                "company_name": name,
                "corporate_number": row.get("corporate_number", ""),
                "entity_type": classify_entity(name),
                "source_queries": set(),
                "regions": set(),
                "office_locations": set(),
                "purposes": set(),
                "license_records": 0,
                "license_date": row.get("license_date", ""),
            })
            entry["license_records"] += 1
            entry["source_queries"].add(row.get("query", ""))
            entry["regions"].add(row.get("region", ""))
            if row.get("office_location"):
                entry["office_locations"].add(row["office_location"])
            if row.get("purpose"):
                entry["purposes"].add(row["purpose"])
            if not entry["corporate_number"] and row.get("corporate_number"):
                entry["corporate_number"] = row["corporate_number"]

    fields = ["company_name", "corporate_number", "entity_type", "source_queries",
              "regions", "office_locations", "purpose", "license_records", "license_date"]
    with MASTER_CSV.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for entry in companies.values():
            writer.writerow({
                "company_name": entry["company_name"],
                "corporate_number": entry["corporate_number"],
                "entity_type": entry["entity_type"],
                "source_queries": ";".join(sorted(q for q in entry["source_queries"] if q)),
                "regions": ";".join(sorted(r for r in entry["regions"] if r)),
                "office_locations": ";".join(sorted(entry["office_locations"])[:3]),
                "purpose": ";".join(sorted(entry["purposes"])[:2]),
                "license_records": entry["license_records"],
                "license_date": entry["license_date"],
            })
    print(f"master: {len(companies)} unique companies -> {MASTER_CSV}")


# ----------------------------------------------------------------- enrich ---

ENRICHED_COLUMNS = FINAL_FRONT_COLUMNS + [
    "source_queries", "regions", "office_locations", "license_records",
    "nta_name", "nta_last_updated",
]


def yen_pretty(value) -> str:
    try:
        amount = int(float(value))
    except (TypeError, ValueError):
        return ""
    if amount >= 100_000_000:
        oku = amount / 100_000_000
        return f"{oku:.1f}億円".replace(".0億円", "億円")
    if amount >= 10_000:
        return f"{amount // 10_000}万円"
    return f"{amount:,}円"


def gbiz_api_prefill(fetcher: ThreadSafeFetcher, token: str, row: dict) -> None:
    number = clean(row.get("corporate_number", ""))
    if not token or not re.fullmatch(r"\d{13}", number):
        return
    headers = {"X-hojinInfo-api-token": token, "Accept": "application/json"}
    base = f"https://info.gbiz.go.jp/hojin/v1/hojin/{number}"
    data = fetcher.fetch_json(base, headers)
    infos = data.get("hojin-infos") or []
    if infos:
        info = infos[0]
        src = f"gBizINFO-API:{number}"
        set_field(row, "detailed_address", info.get("location", ""), src, "gbiz-api")
        set_field(row, "website_url", info.get("company_url", ""), src, "gbiz-api")
        set_field(row, "capital", yen_pretty(info.get("capital_stock")), src, "gbiz-api")
        if info.get("employee_number"):
            set_field(row, "employee_count", f"{info['employee_number']}人", src, "gbiz-api")
        business = " / ".join(filter(None, [info.get("business_summary", ""),
                                            " ".join(info.get("business_items") or [])]))
        set_field(row, "industry_business", business, src, "gbiz-api")
    finance = fetcher.fetch_json(base + "/finance", headers)
    if finance:
        m = re.search(r"(\d{4})年(\d{1,2})月期", json.dumps(finance, ensure_ascii=False))
        if m:
            set_field(row, "fiscal_month", f"{int(m.group(2))}月", f"gBizINFO-API:{number}/finance", "gbiz-api")


def enriched_done_keys() -> set[str]:
    if not ENRICHED_CSV.exists():
        return set()
    keys: set[str] = set()
    with ENRICHED_CSV.open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            keys.add(row.get("corporate_number") or f"name:{clean(row.get('company_name', ''))}")
    return keys


def enrich(workers: int, limit: int, sleep: float, max_pages: int, max_pdfs: int,
           max_pdf_pages: int, search_web: bool) -> None:
    if not MASTER_CSV.exists():
        print("companies_master.csv not found — run collect first")
        return
    with MASTER_CSV.open(encoding="utf-8-sig", newline="") as f:
        master = list(csv.DictReader(f))

    done = enriched_done_keys()
    todo = [r for r in master
            if (r.get("corporate_number") or f"name:{clean(r.get('company_name', ''))}") not in done]
    if limit:
        todo = todo[:limit]
    print(f"enrich: {len(master)} companies, {len(done)} done, {len(todo)} to go, workers={workers}")
    if not todo:
        return

    fetcher = ThreadSafeFetcher(default_interval=sleep)
    token = os.environ.get("GBIZINFO_API_TOKEN", "")
    if token:
        print("gBizINFO API token detected — using REST API as primary structured source")

    new_file = not ENRICHED_CSV.exists()
    out_f = ENRICHED_CSV.open("a", encoding="utf-8-sig", newline="")
    writer = csv.DictWriter(out_f, fieldnames=ENRICHED_COLUMNS, extrasaction="ignore")
    if new_file:
        writer.writeheader()
    write_lock = threading.Lock()
    counter = {"n": 0}

    def work(master_row: dict) -> dict:
        row = ensure_schema(dict(master_row))
        gbiz_api_prefill(fetcher, token, row)
        row = enrich_row(
            fetcher, row,
            max_pages_per_site=max_pages,
            max_pdfs_per_site=max_pdfs,
            max_pdf_pages=max_pdf_pages,
            search_web=search_web,
        )
        return row

    try:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = {pool.submit(work, row): row for row in todo}
            for future in as_completed(futures):
                src = futures[future]
                try:
                    row = future.result()
                except Exception as exc:
                    row = ensure_schema(dict(src))
                    row["enrichment_status"] = "エラー"
                    row["manual_check_reason"] = f"enrich exception: {exc}"
                with write_lock:
                    writer.writerow(row)
                    out_f.flush()
                    counter["n"] += 1
                    n = counter["n"]
                print(f"[{n}/{len(todo)}] {row.get('company_name', '')} "
                      f"web={row.get('website_url', '')} tel={row.get('phone_number', '')} "
                      f"{row.get('manual_check_reason', '')}")
    finally:
        out_f.close()
    print(f"enriched -> {ENRICHED_CSV}")


# ----------------------------------------------------------------- export ---

DELIVERABLE_MAP = [
    ("企業名", "company_name"),
    ("法人番号", "corporate_number"),
    ("区分", "entity_type"),
    ("住所", "detailed_address"),
    ("電話番号", "phone_number"),
    ("WebサイトURL", "website_url"),
    ("決算月", "fiscal_month"),
    ("FAX番号", "fax_number"),
    ("メールアドレス", "email_address"),
    ("問い合わせフォームURL", "contact_form_url"),
    ("業種・事業内容", "industry_business"),
    ("資本金", "capital"),
    ("従業員数", "employee_count"),
    ("抽出元検索", "source_queries"),
    ("地域", "regions"),
    ("免許レコード数", "license_records"),
    ("取得ステータス", "enrichment_status"),
    ("未取得項目", "manual_check_reason"),
    ("住所出典", "address_source_url"),
    ("電話出典", "phone_source_url"),
    ("URL出典", "website_source_url"),
    ("決算月出典", "fiscal_month_source_url"),
    ("FAX出典", "fax_source_url"),
    ("メール出典", "email_source_url"),
    ("フォーム出典", "contact_form_source_url"),
    ("業種出典", "industry_business_source_url"),
    ("資本金出典", "capital_source_url"),
    ("従業員数出典", "employee_count_source_url"),
]


def export() -> None:
    if not ENRICHED_CSV.exists():
        print("enriched_full.csv not found — run enrich first")
        return
    import pandas as pd

    with ENRICHED_CSV.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))
    out_rows = [{ja: row.get(en, "") for ja, en in DELIVERABLE_MAP} for row in rows]

    csv_path = DELIVER_PREFIX.with_suffix(".csv")
    with csv_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=[ja for ja, _ in DELIVERABLE_MAP])
        writer.writeheader()
        writer.writerows(out_rows)

    xlsx_path = DELIVER_PREFIX.with_suffix(".xlsx")
    pd.DataFrame(out_rows).to_excel(xlsx_path, index=False)
    format_xlsx(xlsx_path)
    print(f"deliverable: {len(out_rows)} rows -> {csv_path} / {xlsx_path}")

    fields = ["住所", "電話番号", "WebサイトURL", "決算月", "FAX番号", "メールアドレス",
              "問い合わせフォームURL", "業種・事業内容", "資本金", "従業員数"]
    total = len(out_rows)
    print("\nFill stats")
    for field in fields:
        filled = sum(1 for r in out_rows if r.get(field) and r.get(field) != "対象外")
        na = sum(1 for r in out_rows if r.get(field) == "対象外")
        suffix = f" (+対象外 {na})" if na else ""
        print(f"- {field}: {filled}/{total}{suffix}")


# ------------------------------------------------------------------- main ---

def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("phase", choices=["collect", "enrich", "export", "all"])
    parser.add_argument("--queries", default="q1,q2,q3,q4", help="comma list among q1,q2,q3,q4")
    parser.add_argument("--limit-pages", type=int, default=0, help="max pages per query unit this run (0=all)")
    parser.add_argument("--limit", type=int, default=0, help="max companies to enrich this run (0=all)")
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--sleep", type=float, default=0.35)
    parser.add_argument("--max-pages-per-site", type=int, default=12)
    parser.add_argument("--max-pdfs-per-site", type=int, default=2)
    parser.add_argument("--max-pdf-pages", type=int, default=8)
    parser.add_argument("--no-search-web", action="store_true")
    parser.add_argument("--quick-test", action="store_true",
                        help="smoke run: q3 only, 1 page per unit, enrich 5 companies")
    args = parser.parse_args()

    queries = [q.strip() for q in args.queries.split(",") if q.strip()]
    if args.quick_test:
        queries = ["q3"]
        args.limit_pages = args.limit_pages or 1
        args.limit = args.limit or 5
        args.workers = min(args.workers, 4)
        args.max_pages_per_site = min(args.max_pages_per_site, 8)

    if args.phase in {"collect", "all"}:
        collect(queries, args.limit_pages, max(args.sleep, 0.5))
    if args.phase in {"enrich", "all"}:
        enrich(args.workers, args.limit, args.sleep, args.max_pages_per_site,
               args.max_pdfs_per_site, args.max_pdf_pages, not args.no_search_web)
    if args.phase in {"export", "all"}:
        export()


if __name__ == "__main__":
    main()
