import csv
import re
import time
import urllib.parse
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup


INPUT = Path("tele_soumu_sample_100_final.csv")
OUTPUT = Path("tele_soumu_sample_100_final_plus.csv")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}

COMMON_PATHS = [
    "/contact/",
    "/inquiry/",
    "/otoiawase/",
    "/toiawase/",
    "/company/",
    "/company/outline/",
    "/company/profile/",
    "/corporate/",
    "/corporate/profile/",
    "/about/",
    "/profile/",
]

KEYWORDS = [
    "お問い合わせ",
    "お問合せ",
    "問い合わせ",
    "問合せ",
    "連絡先",
    "組織",
    "会社概要",
    "企業情報",
    "contact",
    "inquiry",
    "profile",
    "company",
    "about",
]


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def normalize(text: str) -> str:
    return text.translate(str.maketrans("０１２３４５６７８９－ー―（）　", "0123456789---() "))


def fetch(url: str) -> tuple[str, str]:
    try:
        res = requests.get(url, headers=HEADERS, timeout=8, allow_redirects=True, verify=False)
        if res.status_code >= 400:
            return "", url
        res.encoding = res.apparent_encoding or res.encoding
        return res.text, res.url
    except Exception:
        return "", url


def good_phone(value: str) -> bool:
    if not value or "-" not in value:
        return False
    if value in {"03-1234-5678"} or value.startswith("000") or value.startswith("0110-"):
        return False
    return bool(re.fullmatch(r"0\d{1,4}-\d{1,4}-\d{3,4}|0\d{2,4}-\d{3,4}", value))


def extract(html: str, source_url: str) -> dict[str, str]:
    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    text = normalize(clean(soup.get_text(" ", strip=True)))
    result = {"phone_number": "", "fax_number": "", "email_address": "", "contact_form_url": ""}

    for match in re.finditer(r"0\d{1,4}[-(]\d{1,4}[-)]\d{3,4}|0\d{9,10}", text):
        number = match.group(0).replace("(", "-").replace(")", "-").strip("-")
        ctx = text[max(0, match.start() - 80) : min(len(text), match.end() + 40)].lower()
        if "fax" in ctx or "ファックス" in ctx:
            if not result["fax_number"] and good_phone(number):
                result["fax_number"] = number
        elif any(word in ctx for word in ["tel", "電話", "代表", "お問い合わせ", "連絡"]):
            if not result["phone_number"] and good_phone(number):
                result["phone_number"] = number

    emails = re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", text)
    if emails:
        result["email_address"] = emails[0]

    for a in soup.find_all("a", href=True):
        label = clean(a.get_text(" ", strip=True)).lower()
        href = a["href"]
        if href.lower().startswith("mailto:") and not result["email_address"]:
            result["email_address"] = href.split(":", 1)[1].split("?", 1)[0]
        if href.lower().startswith("tel:") and not result["phone_number"]:
            candidate = normalize(href.split(":", 1)[1]).strip()
            if good_phone(candidate):
                result["phone_number"] = candidate
        if not result["contact_form_url"]:
            haystack = f"{label} {href.lower()}"
            if any(k.lower() in haystack for k in KEYWORDS[:5]):
                result["contact_form_url"] = urllib.parse.urljoin(source_url, href).split("#")[0]
    return result


def candidates(base_url: str, html: str) -> list[str]:
    parsed = urllib.parse.urlparse(base_url)
    origin = f"{parsed.scheme}://{parsed.netloc}"
    urls = [base_url]
    for path in COMMON_PATHS:
        urls.append(origin + path)
    soup = BeautifulSoup(html, "html.parser")
    for a in soup.find_all("a", href=True):
        label = clean(a.get_text(" ", strip=True)).lower()
        href = a["href"]
        haystack = f"{label} {href.lower()}"
        if any(k.lower() in haystack for k in KEYWORDS):
            u = urllib.parse.urljoin(base_url, href).split("#")[0]
            if urllib.parse.urlparse(u).netloc == parsed.netloc:
                urls.append(u)
    unique = []
    for u in urls:
        if u not in unique:
            unique.append(u)
    return unique[:8]


def main() -> None:
    df = pd.read_csv(INPUT, dtype=str).fillna("")
    for idx, row in df.iterrows():
        if row["phone_number"] and row["contact_form_url"]:
            continue
        website = row["website_url"]
        if not website:
            continue
        html, final_url = fetch(website)
        if not html:
            continue
        found_any = False
        for url in candidates(final_url, html):
            html2, final2 = fetch(url)
            if not html2:
                continue
            data = extract(html2, final2)
            for field, source_field in [
                ("phone_number", "phone_source_url"),
                ("fax_number", "fax_source_url"),
                ("email_address", "email_source_url"),
                ("contact_form_url", "contact_form_source_url"),
            ]:
                if data[field] and not df.at[idx, field]:
                    df.at[idx, field] = data[field]
                    df.at[idx, source_field] = final2
                    found_any = True
            if df.at[idx, "phone_number"] and df.at[idx, "contact_form_url"]:
                break
            time.sleep(0.2)
        print(row["company_name"], "updated" if found_any else "")
        time.sleep(0.3)

    df.to_csv(OUTPUT, encoding="utf-8-sig", index=False, quoting=csv.QUOTE_MINIMAL)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
