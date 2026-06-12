import csv
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
import requests
from bs4 import BeautifulSoup


INPUT = Path("tele_soumu_sample_100_deep_contacts.csv")
OUTPUT = Path("tele_soumu_sample_100_deep_contacts_urls.csv")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
}

OFFICIAL_URLS = {
    "水戸市": "https://www.city.mito.lg.jp/",
    "鉾田市": "https://www.city.hokota.lg.jp/",
    "武州瓦斯株式会社": "https://www.bushugas.co.jp/",
    "蕨市": "https://www.city.warabi.saitama.jp/",
    "株式会社ＺＯＺＯ": "https://corp.zozo.com/",
    "松戸市": "https://www.city.matsudo.chiba.jp/",
    "ジャパンリアルエステイトアセットマネジメント株式会社": "https://www.j-re.co.jp/",
    "ハイウエイ開発株式会社": "https://www.highway-kaihatsu.co.jp/",
    "メルコビルエンジニアリング株式会社": "https://www.melco-buileng.co.jp/",
    "リンク情報システム株式会社": "https://www.lis.co.jp/",
    "外務省": "https://www.mofa.go.jp/mofaj/",
    "学校法人日本大学": "https://www.nihon-u.ac.jp/",
    "株式会社全銀電子債権ネットワーク": "https://www.densai.net/",
    "三菱商事都市開発株式会社": "https://www.mcud.co.jp/",
    "田中電気株式会社": "https://www.tanaka-denki.co.jp/",
    "復興庁": "https://www.reconstruction.go.jp/",
    "ＮＴＴアーバンバリューサポート株式会社": "https://www.ntt-us.com/",
    "楽天カード株式会社": "https://www.rakuten-card.co.jp/",
    "楽天ペイメント株式会社": "https://payment.rakuten.co.jp/",
    "株式会社ＥＮＥＯＳモビリニア": "https://www.eneos-mobilinia.co.jp/",
    "株式会社長谷工ライブネット": "https://www.haseko-hln.co.jp/",
    "地崎道路株式会社": "https://www.chizakidoro.co.jp/",
    "ＳＯＭＰＯホールディングス株式会社": "https://www.sompo-hd.com/",
    "ベル・データ株式会社": "https://www.belldata.com/",
    "一般財団法人移動無線センター": "https://www.mrc.or.jp/",
    "株式会社三越伊勢丹ホールディングス": "https://www.imhds.co.jp/",
    "公益財団法人東京都公園協会": "https://www.tokyo-park.or.jp/",
    "日清食品ホールディングス株式会社": "https://www.nissin.com/jp/",
    "日本ロレアル株式会社": "https://www.loreal.com/ja-jp/japan/",
    "公益財団法人東京動物園協会": "https://www.tokyo-zoo.net/",
    "台東区": "https://www.city.taito.lg.jp/",
    "社会福祉法人同愛記念病院財団": "https://www.doai.jp/",
    "大田区": "https://www.city.ota.tokyo.jp/",
    "渋谷区": "https://www.city.shibuya.tokyo.jp/",
    "全国農業協同組合連合会": "https://www.zennoh.or.jp/",
    "中野区": "https://www.city.tokyo-nakano.lg.jp/",
    "北区": "https://www.city.kita.tokyo.jp/",
}


def fetch_title(url: str) -> tuple[bool, str, str]:
    try:
        res = requests.get(url, headers=HEADERS, timeout=12, allow_redirects=True)
        if res.status_code >= 400:
            return False, url, f"HTTP {res.status_code}"
        res.encoding = res.apparent_encoding or res.encoding
        soup = BeautifulSoup(res.text, "html.parser")
        title = soup.title.get_text(" ", strip=True) if soup.title else ""
        return True, res.url, title
    except Exception as exc:
        return False, url, str(exc)


def normalize_url(url: str) -> str:
    parsed = urlparse(url)
    if parsed.netloc == "www.shinkin.co.jp" and parsed.path.startswith("/chibaskb/"):
        return f"{parsed.scheme}://{parsed.netloc}/chibaskb/"
    return f"{parsed.scheme}://{parsed.netloc}/"


def main() -> None:
    df = pd.read_csv(INPUT, dtype=str).fillna("")
    for idx, row in df.iterrows():
        name = row["company_name"]
        url = OFFICIAL_URLS.get(name)
        if row.get("website_url") and row["website_url"].strip():
            df.at[idx, "website_url"] = normalize_url(row["website_url"])
            continue
        if not url:
            continue
        ok, final_url, title = fetch_title(url)
        print(name, ok, final_url, title[:80])
        if ok:
            df.at[idx, "website_url"] = normalize_url(final_url)
            df.at[idx, "website_source_url"] = final_url

    df.to_csv(OUTPUT, encoding="utf-8-sig", index=False, quoting=csv.QUOTE_MINIMAL)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
