#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主動型 ETF 持股爬蟲
爬取台灣主動型 ETF 每日持股，輸出 Excel

Playwright：Nomura, Capital, Allianz
requests  ：Taishin, First
跳過      ：FuhHwa (00991A), UniPresident (00981A), CTBC (00995A)
"""

import json
import re
import warnings
from datetime import datetime, timezone

import pandas as pd
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright

warnings.filterwarnings("ignore")

OUTPUT_FILE = "etf_holdings.xlsx"
TODAY = datetime.today().strftime("%Y-%m-%d")


# ══════════════════════════════════════════════
# 共用：Playwright 攔截 JSON 回應
# ══════════════════════════════════════════════
def _playwright_fetch(goto_url, api_pattern, accept_cookie_text=None):
    """
    用 Playwright 開啟頁面，攔截 URL 包含 api_pattern 的回應，回傳 JSON。
    accept_cookie_text: 若頁面有 cookie 同意彈窗，傳入按鈕文字。
    """
    result = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1280, "height": 800})

        def on_response(response):
            if api_pattern in response.url and response.status == 200:
                if response.url not in result:
                    try:
                        result[response.url] = response.json()
                    except Exception:
                        pass

        page.on("response", on_response)
        page.goto(goto_url, timeout=30000)
        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_timeout(2000)

        if accept_cookie_text:
            try:
                page.locator(f"text={accept_cookie_text}").click(timeout=3000)
                page.wait_for_timeout(1000)
            except Exception:
                pass

        page.wait_for_timeout(3000)
        browser.close()

    return result


# ══════════════════════════════════════════════
# 1. Nomura 野村投信 — 00980A, 00985A
# ══════════════════════════════════════════════
def fetch_nomura(etf_id: str) -> pd.DataFrame:
    result = _playwright_fetch(
        goto_url=f"https://www.nomurafunds.com.tw/ETF/fund/{etf_id}",
        api_pattern="GetFundAssets",
    )
    if not result:
        raise ValueError(f"Nomura {etf_id}: 未攔截到 GetFundAssets 回應")

    data = next(iter(result.values()))
    rows = data["Entries"]["Data"]["Table"][1]["Rows"]
    try:
        date_str = data["Entries"]["Data"]["Table"][0]["Date"]
    except Exception:
        date_str = TODAY

    records = []
    for row in rows:
        records.append({
            "ETF代號": etf_id,
            "股票代號": row[0],
            "股票名稱": row[1],
            "股數":    row[2],
            "權重(%)": row[3],
            "日期":    date_str,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════
# 2. Capital 群益投信 — 00982A (399), 00992A (500)
# ══════════════════════════════════════════════
CAPITAL_FUND_IDS = {"00982A": "399", "00992A": "500"}

def fetch_capital(etf_id: str) -> pd.DataFrame:
    fund_id = CAPITAL_FUND_IDS[etf_id]
    result = _playwright_fetch(
        goto_url=f"https://www.capitalfund.com.tw/CFWeb/ETF/detail/{fund_id}",
        api_pattern="etf/buyback",
    )
    if not result:
        raise ValueError(f"Capital {etf_id}: 未攔截到 buyback 回應")

    data = next(iter(result.values()))
    date_str = data["data"]["pcf"].get("date2", TODAY)
    records = []
    for s in data["data"]["stocks"]:
        records.append({
            "ETF代號": etf_id,
            "股票代號": s["stocNo"],
            "股票名稱": s["stocName"],
            "股數":    s["share"],
            "權重(%)": s["weight"],
            "日期":    date_str,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════
# 3. Allianz 安聯投信 — 00984A (E0001), 00993A (E0002)
#    需點擊 ETF 連結後才會呼叫 GetFundAssets
# ══════════════════════════════════════════════
ALLIANZ_FUND_IDS = {"00984A": "E0001", "00993A": "E0002"}

def fetch_allianz(etf_id: str) -> pd.DataFrame:
    fund_no = ALLIANZ_FUND_IDS[etf_id]
    result = {}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1280, "height": 800})

        def on_response(response):
            if "GetFundAssets" in response.url and response.status == 200:
                try:
                    result["data"] = response.json()
                except Exception:
                    pass

        page.on("response", on_response)

        # Step 1: 先到 etf-list 觸發 cookie 同意
        page.goto("https://etf.allianzgi.com.tw/etf-list", timeout=30000)
        page.wait_for_load_state("networkidle", timeout=10000)
        page.wait_for_timeout(2000)
        try:
            page.locator("text=接受所有 Cookie").click(timeout=3000)
            page.wait_for_timeout(1000)
        except Exception:
            pass

        # Step 2: 直接到 ETF 持股明細頁（tab=4）
        page.goto(f"https://etf.allianzgi.com.tw/etf-info/{fund_no}?tab=4",
                  timeout=30000)
        page.wait_for_load_state("networkidle", timeout=15000)
        page.wait_for_timeout(5000)

        browser.close()

    if "data" not in result:
        raise ValueError(f"Allianz {etf_id}: 未攔截到 GetFundAssets 回應")

    data = result["data"]
    rows = data["Entries"]["Data"]["Table"][1]["Rows"]
    try:
        # 優先用持股明細的日期（Table[1] 的 Date），再 fallback 到 NavDate
        raw_date = (
            data["Entries"]["Data"]["Table"][1].get("Date")
            or data["Entries"]["Data"]["Table"][0].get("Date")
            or data["Entries"]["Data"]["FundAsset"]["NavDate"]
        )
        date_str = raw_date.replace("/", "-")
    except Exception:
        date_str = TODAY

    records = []
    for row in rows:
        # row: [序號, 股票代號, 股票名稱, 股數, 權重%]
        records.append({
            "ETF代號": etf_id,
            "股票代號": row[1],
            "股票名稱": row[2],
            "股數":    row[3],
            "權重(%)": row[4],
            "日期":    date_str,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════
# 4. Taishin 台新投信 — 00986A, 00987A
#    HTML server-side rendered，持股在最後一個 table
# ══════════════════════════════════════════════
def fetch_taishin(etf_id: str) -> pd.DataFrame:
    url = f"https://www.tsit.com.tw/ETF/Home/ETFSeriesDetail/{etf_id}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    }
    resp = requests.get(url, headers=headers, verify=False, timeout=15)
    resp.raise_for_status()
    resp.encoding = "utf-8"

    soup = BeautifulSoup(resp.text, "html.parser")
    tab = soup.select_one("#paneFive1")
    if not tab:
        raise ValueError(f"找不到 #paneFive1 — {etf_id}")

    # 取公告日期
    date_input = tab.select_one("#PUB_DATE")
    date_str = date_input["value"] if date_input else TODAY

    # 持股在最後一個 table（欄位：代號, 名稱, 股數, 持股比率）
    tables = tab.find_all("table")
    if not tables:
        raise ValueError(f"找不到 table — {etf_id}")
    table = tables[-1]

    records = []
    for row in table.find_all("tr")[1:]:  # 跳過 header
        cols = [td.get_text(strip=True) for td in row.find_all("td")]
        if len(cols) >= 4:
            stock_id = cols[0].split()[0]  # 去除 " TT" 等後綴
            records.append({
                "ETF代號": etf_id,
                "股票代號": stock_id,
                "股票名稱": cols[1],
                "股數":    cols[2],
                "權重(%)": cols[3],
                "日期":    date_str,
            })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════
# 5. First 第一金投信 — 00994A (pStrFundID=182)
# ══════════════════════════════════════════════
def fetch_first(etf_id: str) -> pd.DataFrame:
    url = "https://www.fsitc.com.tw/WebAPI.aspx/Get_hd"
    body = {"pStrFundID": "182", "pStrDate": ""}
    resp = requests.post(url, json=body, verify=False, timeout=15)
    resp.raise_for_status()
    inner = json.loads(resp.json()["d"])

    records = []
    for item in inner:
        if str(item.get("group")) != "1":  # "1"=股票（group 為字串）
            continue
        records.append({
            "ETF代號": etf_id,
            "股票代號": item["A"],
            "股票名稱": item["B"],
            "股數":    item["D"],
            "權重(%)": item["C"],
            "日期":    item.get("sdate", TODAY),
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════
# 6. UniPresident 統一投信 — 00981A (fundCode=49YTW)
#    POST JSON，透過 ezmoney.com.tw 的 GetPCF API
# ══════════════════════════════════════════════
def _dotnet_date_to_str(dotnet_date: str) -> str:
    """將 /Date(1773590400000)/ 轉成 YYYY-MM-DD（台灣時間 UTC+8）"""
    m = re.search(r"/Date\((\d+)\)/", dotnet_date)
    if m:
        from datetime import timedelta
        ts = int(m.group(1)) / 1000
        tw_time = datetime.fromtimestamp(ts, tz=timezone.utc) + timedelta(hours=8)
        return tw_time.strftime("%Y-%m-%d")
    return TODAY

def fetch_uni_president(etf_id: str) -> pd.DataFrame:
    # 取今日民國年日期（ezmoney 用民國年，specificDate=False 拿最新）
    roc_year = datetime.today().year - 1911
    date_str = f"{roc_year}/{datetime.today().strftime('%m/%d')}"

    url = "https://www.ezmoney.com.tw/ETF/Transaction/GetPCF"
    body = {"fundCode": "49YTW", "date": date_str, "specificDate": False}
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://www.ezmoney.com.tw/ETF/Transaction/PCF",
    }
    resp = requests.post(url, json=body, headers=headers, verify=False, timeout=15)
    resp.raise_for_status()
    data = resp.json()

    # 找 AssetCode == "ST" 的那筆，取 Details
    st_asset = next((a for a in data["asset"] if a["AssetCode"] == "ST"), None)
    if not st_asset or not st_asset.get("Details"):
        raise ValueError(f"UniPresident {etf_id}: 找不到股票持股資料")

    tran_date = _dotnet_date_to_str(st_asset["Details"][0]["TranDate"])

    records = []
    for d in st_asset["Details"]:
        records.append({
            "ETF代號": etf_id,
            "股票代號": d["DetailCode"],
            "股票名稱": d["DetailName"],
            "股數":    d["Share"],
            "權重(%)": d["NavRate"],
            "日期":    tran_date,
        })
    return pd.DataFrame(records)


# ══════════════════════════════════════════════
# 跳過（尚未找到可用 API）
#   FuhHwa  00991A — 網址不明
#   CTBC    00995A — 公司防火牆封鎖
# ══════════════════════════════════════════════


# ══════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════
FETCHERS = [
    # 公司防火牆封鎖，需在家裡或其他網路執行
    # ("00980A", fetch_nomura),
    # ("00985A", fetch_nomura),
    # ("00982A", fetch_capital),
    # ("00992A", fetch_capital),
    ("00981A", fetch_uni_president),
    ("00984A", fetch_allianz),
    ("00993A", fetch_allianz),
    ("00986A", fetch_taishin),
    ("00987A", fetch_taishin),
    ("00994A", fetch_first),
]


def main():
    import os

    # 讀取既有資料（若檔案存在）
    existing = {}
    if os.path.exists(OUTPUT_FILE):
        try:
            existing = pd.read_excel(OUTPUT_FILE, sheet_name=None)
            existing.pop("ALL", None)  # ALL 每次重新產生
            print(f"讀取既有檔案：{OUTPUT_FILE}，共 {len(existing)} 個 sheet\n")
        except Exception as e:
            print(f"讀取既有檔案失敗，將建立新檔：{e}\n")

    all_frames = []
    fetched_ids = {etf_id for etf_id, _ in FETCHERS}

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # 本次未爬的 sheet 直接保留寫回
        for sheet_name, df_old in existing.items():
            if sheet_name not in fetched_ids:
                df_old.to_excel(writer, sheet_name=sheet_name, index=False)
                all_frames.append(df_old)

        for etf_id, fetcher in FETCHERS:
            print(f"爬取 {etf_id} ...", end=" ", flush=True)
            try:
                df_new = fetcher(etf_id)
                if etf_id in existing:
                    new_date = df_new["日期"].iloc[0]
                    if new_date in existing[etf_id]["日期"].values:
                        print(f"（日期 {new_date} 已存在，略過）", end=" ")
                        df_combined = existing[etf_id]
                    else:
                        df_combined = pd.concat([existing[etf_id], df_new], ignore_index=True)
                else:
                    df_combined = df_new
                df_combined.to_excel(writer, sheet_name=etf_id, index=False)
                all_frames.append(df_combined)
                print(f"OK（本次 {len(df_new)} 筆，累計 {len(df_combined)} 筆）")
            except Exception as e:
                print(f"FAIL — {e}")
                # 爬取失敗時保留舊資料
                if etf_id in existing:
                    existing[etf_id].to_excel(writer, sheet_name=etf_id, index=False)
                    all_frames.append(existing[etf_id])

        if all_frames:
            pd.concat(all_frames, ignore_index=True).to_excel(
                writer, sheet_name="ALL", index=False
            )

    print(f"\n輸出完成：{OUTPUT_FILE}")


if __name__ == "__main__":
    main()
