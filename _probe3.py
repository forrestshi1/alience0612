"""Probe 3: masked-name ratio across pages; does Q2 detail page reveal names?"""
import sys
import urllib.parse

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, ".")
from bs4 import BeautifulSoup
from tele_soumu_oneclick_enrich import Fetcher, SOUmu_SEARCH, SOUmu_BASE, clean

fetcher = Fetcher(sleep_seconds=0.6)


def fetch(params):
    query = urllib.parse.urlencode(params).replace("+", "%20")
    return fetcher.fetch(f"{SOUmu_SEARCH}?{query}")


def scan(params, label):
    r = fetch(params)
    soup = BeautifulSoup(r.html or "", "html.parser")
    rows = soup.select("tr.m-table-sort__row")
    masked = 0
    corp_samples = []
    for tr in rows:
        name = clean(tr.find("td").get_text(" ", strip=True)) if tr.find("td") else ""
        if name.startswith("*"):
            masked += 1
        elif len(corp_samples) < 4:
            corp_samples.append(name[:70])
    print(f"{label}: rows={len(rows)} masked={masked}")
    for s in corp_samples:
        print("   corp:", s)


# Q1 nara at different offsets
for sc in ["1", "1001", "2001", "2701"]:
    scan({"SK": "2", "DC": "100", "SC": sc, "pageID": "3", "CONFIRM": "1",
          "SelectID": "1", "SelectOW": "03", "HC": "29"}, f"Q1 nara SC={sc}")

# Q2 tokai at different offsets
for sc in ["1", "4001", "9001"]:
    scan({"SC": sc, "pageID": "3", "SelectID": "6", "CONFIRM": "1",
          "OW2": "008", "IT": "C", "DC": "100"}, f"Q2 tokai SC={sc}")

# Q2 detail page: does it reveal the registrant name?
detail = fetcher.fetch(SOUmu_BASE + "SearchServlet?pageID=4&IT=C&DFCD=0000040517&DD=3&styleNumber=")
soup = BeautifulSoup(detail.html or "", "html.parser")
print("\nQ2 detail rows:")
for tr in soup.find_all("tr")[:10]:
    cells = [clean(c.get_text(" ", strip=True))[:60] for c in tr.find_all(["th", "td"])]
    if cells:
        print("  ", cells)
