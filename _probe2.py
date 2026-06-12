"""Probe 2: find real pagination param; sample corporate (non-masked) rows in Q1/Q2."""
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


Q3 = {"SC": "1", "pageID": "3", "SelectID": "5", "CONFIRM": "1", "OW": "ML 1",
      "IT": "A", "FF": "897.5", "TF": "897.5", "HZ": "3", "DC": "100"}

# --- find pager links on page 1 ---
r1 = fetch(Q3)
soup = BeautifulSoup(r1.html, "html.parser")
pager_links = set()
for a in soup.find_all("a", href=True):
    href = a["href"]
    if "SearchServlet" in href and ("SC=" in href or "page" in href.lower()):
        pager_links.add(href[:160])
for link in sorted(pager_links)[:12]:
    print("PAGER:", link)

# also look for forms/hidden inputs that drive paging
for form in soup.find_all("form"):
    hidden = [(i.get("name"), i.get("value")) for i in form.find_all("input", {"type": "hidden"})]
    if hidden:
        print("FORM action=", form.get("action"), "hidden=", hidden[:12])

first_name = clean(soup.select_one("tr.m-table-sort__row td").get_text(" ", strip=True))
print("page1 first:", first_name[:50])

# --- try SC=101 (record offset) ---
r = fetch({**Q3, "SC": "101"})
s2 = BeautifulSoup(r.html, "html.parser")
rows2 = s2.select("tr.m-table-sort__row")
if rows2:
    print(f"SC=101: rows={len(rows2)} first:", clean(rows2[0].find("td").get_text(' ', strip=True))[:50])
else:
    print("SC=101: no rows")

# --- Q1 nara: count masked vs corporate rows, show 3 corporate rows ---
r = fetch({"SK": "2", "DC": "100", "SC": "1", "pageID": "3", "CONFIRM": "1",
           "SelectID": "1", "SelectOW": "03", "HC": "29"})
soup = BeautifulSoup(r.html, "html.parser")
rows = soup.select("tr.m-table-sort__row")
masked = sum(1 for tr in rows if "*****" in tr.get_text())
print(f"\nQ1 nara page1: rows={len(rows)} masked={masked}")
shown = 0
for tr in rows:
    if "*****" in tr.get_text():
        continue
    cells = [clean(td.get_text(" ", strip=True))[:70] for td in tr.find_all("td")]
    print("  corp row:", cells)
    shown += 1
    if shown >= 3:
        break

# --- Q2 tokai: same ---
r = fetch({"SC": "1", "pageID": "3", "SelectID": "6", "CONFIRM": "1",
           "OW2": "008", "IT": "C", "DC": "100"})
soup = BeautifulSoup(r.html, "html.parser")
rows = soup.select("tr.m-table-sort__row")
masked = sum(1 for tr in rows if "*****" in tr.get_text())
print(f"\nQ2 tokai page1: rows={len(rows)} masked={masked}")
shown = 0
for tr in rows:
    if "*****" in tr.get_text():
        continue
    cells = [clean(td.get_text(" ", strip=True))[:70] for td in tr.find_all("td")]
    links = [a.get("href", "")[:90] for a in tr.find_all("a", href=True)]
    print("  corp row:", cells)
    print("  link:", links)
    shown += 1
    if shown >= 3:
        break
