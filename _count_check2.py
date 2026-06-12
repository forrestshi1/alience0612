import re
import sys
import urllib.parse

sys.path.insert(0, ".")
from tele_soumu_oneclick_enrich import Fetcher, SOUmu_SEARCH

fetcher = Fetcher(sleep_seconds=0.6)

# --- Q1 counts for more prefectures (HC = JIS prefecture code) ---
prefs = {"14 kanagawa": "14", "23 aichi": "23", "27 osaka": "27"}
for label, code in prefs.items():
    params = {"SK": "2", "DC": "100", "SC": "1", "pageID": "3", "CONFIRM": "1",
              "SelectID": "1", "SelectOW": "03", "HC": code}
    query = urllib.parse.urlencode(params).replace("+", "%20")
    r = fetcher.fetch(f"{SOUmu_SEARCH}?{query}")
    m = re.search(r"([\d,]+)\s*件中", r.html or "")
    print(f"Q1 {label}: total={m.group(1) if m else '?'}")

# --- Q2: discover form params for 登録情報検索 (SelectID=6) ---
form = fetcher.fetch(f"{SOUmu_SEARCH}?pageID=2&SelectID=6")
html = form.html or ""
print("form status:", form.status_code, "error-page:" , "Error Page" in html)
for m in re.finditer(r'<select[^>]*name="([^"]+)"[^>]*>(.*?)</select>', html, re.S):
    name, body = m.group(1), m.group(2)
    opts = re.findall(r'<option[^>]*value="([^"]*)"[^>]*>\s*([^<]{0,30})', body)
    print(f"SELECT {name}: {opts[:8]}")
for m in re.finditer(r'<input[^>]*name="([^"]+)"[^>]*value="([^"]*)"[^>]*>', html):
    print("INPUT", m.group(1), "=", m.group(2)[:30])
