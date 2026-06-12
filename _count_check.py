"""Quick volume check for the four client queries (counts only, no detail pages)."""
import re
import sys
import urllib.parse

sys.path.insert(0, ".")
from tele_soumu_oneclick_enrich import Fetcher, SOUmu_SEARCH

fetcher = Fetcher(sleep_seconds=0.5)


def count_for(params: dict[str, str], label: str) -> None:
    query = urllib.parse.urlencode(params).replace("+", "%20")
    url = f"{SOUmu_SEARCH}?{query}"
    result = fetcher.fetch(url)
    html = result.html or ""
    rows = html.count("m-table-sort__row")
    m = re.search(r"([\d,]+)\s*件中", html)
    m2 = re.findall(r"全\s*([\d,]+)\s*件|該当[^0-9<]{0,10}([\d,]+)\s*件|([\d,]+)\s*件が該当", html)
    total = m.group(1) if m else (m2[0] if m2 else "?")
    err = "ERROR-PAGE" if "Error Page" in html else ""
    print(f"{label}: rows_on_page={rows} total={total} status={result.status_code} {err} {result.error}")
    if rows == 0 and not err:
        snippet = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html))[:300]
        print("  snippet:", snippet)


# Q3: 陸上移動局(包括免許) 897.5MHz, Kanto (IT=A) — known-good from existing code
count_for({"SC": "1", "pageID": "3", "SelectID": "5", "CONFIRM": "1", "OW": "ML 1",
           "IT": "A", "FF": "897.5", "TF": "897.5", "HZ": "3", "DC": "100"}, "Q3 kanto 897.5MHz")

# Q4: same without frequency filter, Kanto
count_for({"SC": "1", "pageID": "3", "SelectID": "5", "CONFIRM": "1", "OW": "ML 1",
           "IT": "A", "HZ": "3", "DC": "100"}, "Q4 kanto all-freq")

# Q1: 簡易無線局 simple search, Tokyo (HC=都道府県?) — use client's URL params
count_for({"SK": "2", "DC": "100", "SC": "1", "pageID": "3", "CONFIRM": "1",
           "SelectID": "1", "SelectOW": "03", "HC": "13"}, "Q1 tokyo kan-i")

# Q2: エキスパート検索(登録局) デジタル簡易無線局, Kanto
count_for({"pageID": "3", "SelectID": "6", "CONFIRM": "1", "OW": "DR", "IT": "A", "DC": "100"},
          "Q2 kanto digital-kan-i-touroku")
