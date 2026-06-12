import csv
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd


INPUT = Path("tele_soumu_sample_100_deep_contacts_urls.csv")
OUTPUT = Path("tele_soumu_sample_100_manual.csv")


VERIFIED = {
    "ハイウエイ開発株式会社": {
        "website_url": "https://highway-k.subaru-kougyou.jp/",
        "phone_number": "03-3528-8254",
        "fax_number": "03-3528-8147",
        "contact_form_url": "https://highway-k.subaru-kougyou.jp/contact/",
        "capital": "100,000,000円",
        "employee_count": "136名（パート・アルバイト含む）［2026年1月現在］",
        "industry_business": "道路の清掃、植栽、維持修繕、構造物及び道路付属物の補修工事、道路交通管理、交通誘導警備、パーキングエリアにおけるハイウェイショップ経営",
        "source": "https://highway-k.subaru-kougyou.jp/company/index.html",
    },
    "メルコビルエンジニアリング株式会社": {
        "website_url": "https://www.resco.co.jp/",
        "phone_number": "03-6257-8931",
        "fax_number": "03-6257-8939",
        "contact_form_url": "https://www.resco.co.jp/ryodenlift/contact.html",
        "capital": "2億円",
        "employee_count": "約1,300名（2025年10月）",
        "industry_business": "三菱昇降機の販売・設計・工事、小荷物専用昇降機の販売・設計・工事・保守",
        "source": "https://www.resco.co.jp/",
    },
    "株式会社ＥＮＥＯＳモビリニア": {
        "website_url": "https://www.eneos-mobilineer.com/",
        "phone_number": "03-6435-8911",
        "contact_form_url": "https://f.msgs.jp/webapp/form/11311_btq_44/index.do",
        "capital": "1億円",
        "employee_count": "約3,000名（2026年4月1日現在）",
        "industry_business": "サービスステーション（SS）の運営",
        "source": "https://www.eneos-mobilineer.com/about/corporate/",
    },
    "地崎道路株式会社": {
        "website_url": "https://www.chizakiroad.co.jp/",
        "phone_number": "03-5460-1031",
        "fax_number": "03-5460-1036",
        "email_address": "info@chizakiroad.co.jp",
        "contact_form_url": "https://www.chizakiroad.co.jp/information1/contact/",
        "capital": "3億5千万円（2025年3月末現在）",
        "employee_count": "143名（2026年4月1日現在）",
        "industry_business": "道路・施設舗装工事、一般土木工事、生活関連工事、航空機着陸拘束装置、空港メンテナンス工事、油汚染浄化事業、工事資材製造販売等",
        "source": "https://www.chizakiroad.co.jp/company/outline/",
    },
    "公益財団法人東京都公園協会": {
        "website_url": "https://www.tokyo-park.or.jp/",
        "contact_form_url": "https://www.tokyo-park.or.jp/inquiry/",
        "industry_business": "都立公園・庭園・植物園、霊園、水上バス・河川事業、緑と水の市民カレッジ等",
        "source": "https://www.tokyo-park.or.jp/",
    },
    "ＮＴＴアーバンバリューサポート株式会社": {
        "website_url": "https://www.ntt-uvs.com/",
        "source": "https://www.ntt-uvs.com/",
    },
}


def normalize_site(url: str) -> str:
    if not url:
        return ""
    parsed = urlparse(url)
    if not parsed.scheme or not parsed.netloc:
        return url
    if parsed.netloc == "www.shinkin.co.jp" and parsed.path.startswith("/chibaskb/"):
        return f"{parsed.scheme}://{parsed.netloc}/chibaskb/"
    return f"{parsed.scheme}://{parsed.netloc}/"


def set_value(df: pd.DataFrame, idx: int, field: str, value: str, source: str) -> None:
    if not value:
        return
    df.at[idx, field] = value
    source_field = {
        "website_url": "website_source_url",
        "phone_number": "phone_source_url",
        "fax_number": "fax_source_url",
        "email_address": "email_source_url",
        "contact_form_url": "contact_form_source_url",
        "capital": "capital_source_url",
        "employee_count": "employee_count_source_url",
        "industry_business": "industry_business_source_url",
    }.get(field)
    if source_field:
        df.at[idx, source_field] = source


def main() -> None:
    df = pd.read_csv(INPUT, dtype=str).fillna("")
    for idx, row in df.iterrows():
        if row["website_url"]:
            df.at[idx, "website_url"] = normalize_site(row["website_url"])
        data = VERIFIED.get(row["company_name"])
        if not data:
            continue
        source = data["source"]
        for field, value in data.items():
            if field == "source":
                continue
            set_value(df, idx, field, value, source)
        df.at[idx, "website_url"] = normalize_site(df.at[idx, "website_url"])
        print(row["company_name"], source)

    df.to_csv(OUTPUT, encoding="utf-8-sig", index=False, quoting=csv.QUOTE_MINIMAL)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
