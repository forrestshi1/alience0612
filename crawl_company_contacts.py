import csv
import re
import time
import urllib.parse
from pathlib import Path

import requests
from bs4 import BeautifulSoup


INPUT = Path("tele_soumu_sample_100_manual.csv")
OUTPUT = Path("tele_soumu_sample_100_final_raw.csv")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0 Safari/537.36"
    ),
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}

CONTACT_WORDS = [
    "お問い合わせ",
    "お問合せ",
    "問い合わせ",
    "問合せ",
    "contact",
    "inquiry",
    "otoiawase",
]

INFO_WORDS = [
    "会社概要",
    "企業情報",
    "会社情報",
    "概要",
    "アクセス",
    "拠点",
    "profile",
    "company",
    "corporate",
    "about",
    "access",
]


def normalize(text: str) -> str:
    trans = str.maketrans("０１２３４５６７８９－ー―（）　", "0123456789---() ")
    return text.translate(trans)


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def fetch(url: str) -> tuple[str, str]:
    try:
        res = requests.get(url, headers=HEADERS, timeout=12, allow_redirects=True)
        res.raise_for_status()
        res.encoding = res.apparent_encoding or res.encoding
        return res.text, res.url
    except Exception:
        return "", url


def same_site(base: str, target: str) -> bool:
    b = urllib.parse.urlparse(base)
    t = urllib.parse.urlparse(target)
    return t.netloc == "" or t.netloc == b.netloc or t.netloc.endswith("." + b.netloc)


def candidate_links(base_url: str, soup: BeautifulSoup) -> list[str]:
    links: list[str] = []
    for a in soup.find_all("a", href=True):
        label = clean(a.get_text(" ", strip=True)).lower()
        href = a["href"]
        href_l = href.lower()
        if href_l.startswith(("mailto:", "tel:", "javascript:")):
            continue
        url = urllib.parse.urljoin(base_url, href)
        if not same_site(base_url, url):
            continue
        haystack = f"{label} {href_l}"
        if any(word.lower() in haystack for word in CONTACT_WORDS + INFO_WORDS):
            links.append(url.split("#")[0])
    unique: list[str] = []
    for link in links:
        if link not in unique:
            unique.append(link)
    return unique[:8]


def extract_contacts(html: str, url: str) -> dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = normalize(clean(soup.get_text(" ", strip=True)))
    result = {
        "phone_number": "",
        "fax_number": "",
        "email_address": "",
        "contact_form_url": "",
    }

    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    if emails:
        result["email_address"] = emails[0]

    phones = []
    for match in re.finditer(r"0\d{1,4}[-(]\d{1,4}[-)]\d{3,4}|0\d{9,10}", text):
        number = match.group(0).replace("(", "-").replace(")", "-")
        start = max(0, match.start() - 60)
        end = min(len(text), match.end() + 30)
        ctx = text[start:end].lower()
        phones.append((number, ctx))
    for number, ctx in phones:
        if not result["fax_number"] and ("fax" in ctx or "ファックス" in ctx):
            result["fax_number"] = number
        if not result["phone_number"] and ("tel" in ctx or "電話" in ctx or "代表" in ctx):
            result["phone_number"] = number
    if not result["phone_number"] and phones:
        result["phone_number"] = phones[0][0]

    for a in soup.find_all("a", href=True):
        label = clean(a.get_text(" ", strip=True)).lower()
        href = a["href"]
        if href.lower().startswith("mailto:") and not result["email_address"]:
            result["email_address"] = href.split(":", 1)[1].split("?", 1)[0]
        if href.lower().startswith("tel:") and not result["phone_number"]:
            result["phone_number"] = normalize(href.split(":", 1)[1])
        haystack = f"{label} {href.lower()}"
        if not result["contact_form_url"] and any(word.lower() in haystack for word in CONTACT_WORDS):
            result["contact_form_url"] = urllib.parse.urljoin(url, href).split("#")[0]
    return result


def merge_contact(old: dict[str, str], new: dict[str, str], source: str) -> None:
    source_map = {
        "phone_number": "phone_source_url",
        "fax_number": "fax_source_url",
        "email_address": "email_source_url",
        "contact_form_url": "contact_form_source_url",
    }
    for key, value in new.items():
        if value and not old.get(key):
            old[key] = value
            old[source_map[key]] = source


def crawl_row(row: dict[str, str]) -> dict[str, str]:
    start_url = row.get("website_url", "").strip()
    if not start_url:
        return row
    if not start_url.startswith(("http://", "https://")):
        start_url = "https://" + start_url
    html, final_url = fetch(start_url)
    if not html:
        return row
    soup = BeautifulSoup(html, "html.parser")
    merge_contact(row, extract_contacts(html, final_url), final_url)
    for link in candidate_links(final_url, soup):
        if row.get("phone_number") and row.get("fax_number") and row.get("email_address") and row.get("contact_form_url"):
            break
        html2, final2 = fetch(link)
        if not html2:
            continue
        merge_contact(row, extract_contacts(html2, final2), final2)
        time.sleep(0.4)
    return row


def main() -> None:
    with INPUT.open(encoding="utf-8-sig", newline="") as f:
        rows = list(csv.DictReader(f))

    out_rows = []
    for i, row in enumerate(rows, 1):
        row = crawl_row(row)
        out_rows.append(row)
        print(
            f"{i:03d} {row['company_name']} "
            f"TEL={row.get('phone_number','')} FAX={row.get('fax_number','')} "
            f"MAIL={row.get('email_address','')} FORM={row.get('contact_form_url','')}"
        )
        time.sleep(0.6)

    with OUTPUT.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(out_rows[0].keys()))
        writer.writeheader()
        writer.writerows(out_rows)
    print(f"Wrote {len(out_rows)} rows to {OUTPUT}")


if __name__ == "__main__":
    main()
