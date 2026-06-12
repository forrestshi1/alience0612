import csv
import re
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass

from bs4 import BeautifulSoup


BASE = "https://www.tele.soumu.go.jp/musen/"
SEARCH = urllib.parse.urljoin(BASE, "SearchServlet")
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


@dataclass
class SearchConfig:
    name: str
    params: dict[str, str]


def fetch(url: str, retries: int = 3) -> str:
    last_error: Exception | None = None
    for attempt in range(retries):
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=30) as res:
                return res.read().decode("utf-8", errors="replace")
        except Exception as exc:
            last_error = exc
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"Failed to fetch {url}") from last_error


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def build_url(params: dict[str, str]) -> str:
    query = urllib.parse.urlencode(params).replace("+", "%20")
    return f"{SEARCH}?{query}#result"


def parse_name_and_corporate_number(raw_name: str) -> tuple[str, str]:
    match = re.search(r"^(.*?)\s+.*?(\d{13})$", raw_name)
    if not match:
        return raw_name, ""
    return match.group(1).strip(), match.group(2)


def table_rows(soup: BeautifulSoup) -> list[list[str]]:
    rows: list[list[str]] = []
    table = soup.find("table")
    if not table:
        return rows
    for tr in table.find_all("tr"):
        cells = [clean(cell.get_text(" ", strip=True)) for cell in tr.find_all(["th", "td"])]
        if cells:
            rows.append(cells)
    return rows


def parse_detail(url: str) -> dict[str, str]:
    soup = BeautifulSoup(fetch(url), "html.parser")
    rows = table_rows(soup)

    def cell(row_index: int, col_index: int = 1) -> str:
        if row_index >= len(rows):
            return ""
        row = rows[row_index]
        return row[col_index] if len(row) > col_index else ""

    return {
        "detail_licensee_name": cell(0),
        "detail_address": cell(1),
        "station_type": cell(2),
        "license_number": cell(3, 3),
        "valid_period": cell(4, 3),
        "station_count": cell(5),
        "office_location_detail": cell(7),
    }


def parse_list(config: SearchConfig, limit: int) -> list[dict[str, str]]:
    soup = BeautifulSoup(fetch(build_url(config.params)), "html.parser")
    records: list[dict[str, str]] = []
    for tr in soup.select("tr.m-table-sort__row"):
        cells = [clean(cell.get_text(" ", strip=True)) for cell in tr.find_all("td")]
        link = tr.select_one('a[href*="pageID=4"]')
        if len(cells) < 4 or not link:
            continue
        detail_url = urllib.parse.urljoin(BASE, link["href"])
        company_name, corporate_number = parse_name_and_corporate_number(cells[0])
        detail = parse_detail(detail_url)
        records.append(
            {
                "source_query": config.name,
                "company_name": company_name,
                "corporate_number": corporate_number,
                "office_location": cells[1],
                "purpose": cells[2],
                "license_date": cells[3],
                "detail_url": detail_url,
                **detail,
            }
        )
        if len(records) >= limit:
            break
        time.sleep(0.7)
    return records


def main() -> None:
    config = SearchConfig(
        name="q3_kanto_land_mobile_blanket_897_5mhz",
        params={
            "SC": "1",
            "pageID": "3",
            "SelectID": "5",
            "CONFIRM": "1",
            "OW": "ML 1",
            "IT": "A",
            "FF": "897.5",
            "TF": "897.5",
            "HZ": "3",
            "DC": "100",
        },
    )
    records = parse_list(config, limit=100)
    output = "tele_soumu_sample_100.csv"
    fieldnames = [
        "source_query",
        "company_name",
        "detail_licensee_name",
        "corporate_number",
        "office_location",
        "purpose",
        "license_date",
        "station_type",
        "station_count",
        "office_location_detail",
        "detail_address",
        "valid_period",
        "license_number",
        "detail_url",
    ]
    with open(output, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(records)
    print(f"Wrote {len(records)} rows to {output}")


if __name__ == "__main__":
    main()
