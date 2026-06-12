import re
import sys
import urllib.parse

sys.path.insert(0, ".")
from tele_soumu_oneclick_enrich import Fetcher, SOUmu_SEARCH

fetcher = Fetcher(sleep_seconds=0.6)


def total(params: dict[str, str]) -> str:
    query = urllib.parse.urlencode(params).replace("+", "%20")
    r = fetcher.fetch(f"{SOUmu_SEARCH}?{query}")
    m = re.search(r"([\d,]+)\s*件中", r.html or "")
    return m.group(1) if m else "?"


# Q1 remaining prefectures
for label, code in [("saitama", "11"), ("chiba", "12"), ("gunma", "10"), ("ibaraki", "8"),
                    ("shizuoka", "22"), ("kyoto", "26"), ("nara", "29")]:
    print(f"Q1 {label}: {total({'SK': '2', 'DC': '100', 'SC': '1', 'pageID': '3', 'CONFIRM': '1', 'SelectID': '1', 'SelectOW': '03', 'HC': code})}")

# Q2 digital kan-i registered stations per bureau
for label, it in [("kanto", "A"), ("tokai", "C"), ("kinki", "E")]:
    print(f"Q2 {label}: {total({'SC': '1', 'pageID': '3', 'SelectID': '6', 'CONFIRM': '1', 'OW2': '008', 'IT': it, 'DC': '100'})}")

# Q3 897.5MHz per bureau
for label, it in [("tokai", "C"), ("kinki", "E")]:
    print(f"Q3 {label}: {total({'SC': '1', 'pageID': '3', 'SelectID': '5', 'CONFIRM': '1', 'OW': 'ML 1', 'IT': it, 'FF': '897.5', 'TF': '897.5', 'HZ': '3', 'DC': '100'})}")

# Q4 all-freq land-mobile blanket per bureau
for label, it in [("tokai", "C"), ("kinki", "E")]:
    print(f"Q4 {label}: {total({'SC': '1', 'pageID': '3', 'SelectID': '5', 'CONFIRM': '1', 'OW': 'ML 1', 'IT': it, 'HZ': '3', 'DC': '100'})}")
