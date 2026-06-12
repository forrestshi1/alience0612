"""Probe pagination params and list-row structure for all 4 query types."""
import re
import sys
import urllib.parse

sys.path.insert(0, ".")
from bs4 import BeautifulSoup
from tele_soumu_oneclick_enrich import Fetcher, SOUmu_SEARCH, clean

fetcher = Fetcher(sleep_seconds=0.6)


def fetch(params):
    query = urllib.parse.urlencode(params).replace("+", "%20")
    return fetcher.fetch(f"{SOUmu_SEARCH}?{query}")


def dump_rows(html, label, max_rows=3):
    soup = BeautifulSoup(html or "", "html.parser")
    rows = soup.select("tr.m-table-sort__row")
    total = re.search(r"([\d,]+)\s*件中", html or "")
    print(f"\n=== {label}: rows={len(rows)} total={total.group(1) if total else '?'}")
    # header row
    table = soup.find("table")
    if table:
        ths = [clean(th.get_text(' ', strip=True)) for th in table.find_all("th")][:8]
        print("  headers:", ths)
    for tr in rows[:max_rows]:
        cells = [clean(td.get_text(" ", strip=True))[:60] for td in tr.find_all("td")]
        links = [a.get("href", "")[:90] for a in tr.find_all("a", href=True)]
        print("  cells:", cells)
        print("  links:", links)


Q3 = {"SC": "1", "pageID": "3", "SelectID": "5", "CONFIRM": "1", "OW": "ML 1",
      "IT": "A", "FF": "897.5", "TF": "897.5", "HZ": "3", "DC": "100"}

# --- pagination: Q3 kanto has 123 items. Try SC=2 (page index) ---
r1 = fetch(Q3)
dump_rows(r1.html, "Q3 page SC=1", 1)
r2 = fetch({**Q3, "SC": "2"})
dump_rows(r2.html, "Q3 page SC=2", 2)

# --- Q1 簡易無線局 (SelectID=1), Nara (small: 2779) ---
r = fetch({"SK": "2", "DC": "100", "SC": "1", "pageID": "3", "CONFIRM": "1",
           "SelectID": "1", "SelectOW": "03", "HC": "29"})
dump_rows(r.html, "Q1 nara SelectID=1")

# --- Q2 登録局 (SelectID=6), Tokai (9467) ---
r = fetch({"SC": "1", "pageID": "3", "SelectID": "6", "CONFIRM": "1",
           "OW2": "008", "IT": "C", "DC": "100"})
dump_rows(r.html, "Q2 tokai SelectID=6")
