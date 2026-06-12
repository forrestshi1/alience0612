import csv
import re
import time
import urllib.request
from pathlib import Path

from bs4 import BeautifulSoup


INPUT = Path("tele_soumu_sample_100.csv")
OUTPUT = Path("tele_soumu_sample_100_enriched.csv")
NTA_URL = "https://www.houjin-bangou.nta.go.jp/henkorireki-johoto.html?selHouzinNo={}"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}


EXTRA_CORPORATE_NUMBERS = {
    "株式会社Ｌ＆Ｆアセットファイナンス": "9010001060224",
    "三井住友トラスト・パナソニックファイナンス株式会社": "1010001146146",
    "株式会社三越伊勢丹ホールディングス": "3011101060499",
    "東急不動産株式会社": "7011001016580",
}

GBIZ_URL = "https://info.gbiz.go.jp/hojin/ichiran?hojinBango={}"


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


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


def parse_nta_profile(corporate_number: str) -> dict[str, str]:
    if not corporate_number:
        return {}
    soup = BeautifulSoup(fetch(NTA_URL.format(corporate_number)), "html.parser")
    profile: dict[str, str] = {}
    dts = soup.find_all("dt")
    for dt in dts:
        dd = dt.find_next_sibling("dd")
        if not dd:
            continue
        key = clean(dt.get_text(" ", strip=True))
        value = clean(dd.get_text(" ", strip=True))
        profile[key] = value
    return {
        "nta_name": profile.get("商号又は名称", ""),
        "detailed_address": profile.get("本店又は主たる事務所の所在地", ""),
        "nta_last_updated": profile.get("最終更新年月日", ""),
        "nta_source_url": NTA_URL.format(corporate_number),
    }


def strip_source(text: str) -> str:
    text = re.sub(r"\s*[（(]\s*[^（）()]*?(法人番号公表サイト|職場情報総合サイト|EDINET|GEPS)[^（）()]*?\s*[）)]", "", text)
    return clean(text)


def parse_gbiz_profile(corporate_number: str) -> dict[str, str]:
    if not corporate_number:
        return {}
    url = GBIZ_URL.format(corporate_number)
    soup = BeautifulSoup(fetch(url), "html.parser")
    data: dict[str, str] = {"gbiz_source_url": url}

    first_body = soup.select_one(".accordion-body")
    if first_body:
        for item in first_body.select(".row.mt-3"):
            children = item.find_all(recursive=False)
            if len(children) < 2:
                continue
            key = clean(children[0].get_text(" ", strip=True))
            value = clean(children[1].get_text(" ", strip=True))
            if key and value:
                data[key] = strip_source(value)

    text = clean(soup.get_text(" ", strip=True))
    match = re.search(r"当期\s+第[^（]*（自\s+\d{4}年\d{1,2}月\d{1,2}日\s+至\s+\d{4}年(\d{1,2})月\d{1,2}日", text)
    if match:
        data["決算月"] = f"{match.group(1)}月"
    return data


def classify_entity(name: str, corporate_number: str) -> str:
    if re.search(r"(市|区|町|村)$", name):
        return "自治体"
    if name.endswith("省") or name.endswith("庁"):
        return "官公庁"
    if "学校法人" in name:
        return "学校法人"
    if "社会福祉法人" in name:
        return "社会福祉法人"
    if "公益財団法人" in name or "一般財団法人" in name:
        return "財団法人"
    if "生活協同組合" in name or "信用金庫" in name or name.endswith("連合会"):
        return "その他法人"
    return "株式会社等"


def not_applicable_reason(entity_type: str, field: str) -> str:
    if entity_type in {"自治体", "官公庁"} and field in {"capital", "fiscal_month"}:
        return "対象外"
    return ""


def main() -> None:
    with INPUT.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    out_rows: list[dict[str, str]] = []
    for index, row in enumerate(rows, 1):
        corporate_number = row.get("corporate_number", "").strip()
        if not corporate_number:
            corporate_number = EXTRA_CORPORATE_NUMBERS.get(row["company_name"], "")

        profile = parse_nta_profile(corporate_number)
        gbiz = parse_gbiz_profile(corporate_number)
        entity_type = classify_entity(row["company_name"], corporate_number)
        business = gbiz.get("業種", "")
        if gbiz.get("事業概要"):
            business = f"{business} / {gbiz['事業概要']}" if business else gbiz["事業概要"]

        out = {
            **row,
            "corporate_number": corporate_number,
            "entity_type": entity_type,
            "detailed_address": profile.get("detailed_address", "") or gbiz.get("本店所在地", ""),
            "address_source": "国税庁法人番号公表サイト" if profile.get("detailed_address") else ("Gビズインフォ" if gbiz.get("本店所在地") else ""),
            "nta_name": profile.get("nta_name", ""),
            "nta_last_updated": profile.get("nta_last_updated", ""),
            "nta_source_url": profile.get("nta_source_url", ""),
            "phone_number": "",
            "phone_source_url": "",
            "website_url": gbiz.get("企業ホームページ", ""),
            "website_source_url": gbiz.get("gbiz_source_url", "") if gbiz.get("企業ホームページ") else "",
            "fiscal_month": not_applicable_reason(entity_type, "fiscal_month") or gbiz.get("決算月", ""),
            "fiscal_month_source_url": gbiz.get("gbiz_source_url", "") if gbiz.get("決算月", "") else "",
            "fax_number": "",
            "fax_source_url": "",
            "email_address": "",
            "email_source_url": "",
            "contact_form_url": "",
            "contact_form_source_url": "",
            "industry_business": business,
            "industry_business_source_url": gbiz.get("gbiz_source_url", "") if business else "",
            "capital": not_applicable_reason(entity_type, "capital") or gbiz.get("資本金", ""),
            "capital_source_url": gbiz.get("gbiz_source_url", "") if gbiz.get("資本金") else "",
            "employee_count": gbiz.get("従業員数", ""),
            "employee_count_source_url": gbiz.get("gbiz_source_url", "") if gbiz.get("従業員数") else "",
            "enrichment_status": (
                "公的DB補完済み・電話/FAX/メール/問い合わせフォームは公式サイト調査が必要"
                if profile.get("detailed_address")
                else "法人番号未確定・追加調査が必要"
            ),
        }
        out_rows.append(out)
        print(f"{index:03d} {row['company_name']} {out['detailed_address']}")
        time.sleep(0.35)

    fieldnames = list(out_rows[0].keys())
    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(out_rows)
    print(f"Wrote {len(out_rows)} rows to {OUTPUT}")


if __name__ == "__main__":
    main()
