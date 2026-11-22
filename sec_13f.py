# -*- coding: utf-8 -*-
"""
미국 SEC EDGAR 공식 API 기반 헤지펀드 13F 공시 자동 추적 모듈 (무료 공식 데이터)
- 추적 대상 헤지펀드(Citadel, Millennium, Point72, Jane Street, SIG, Bridgewater, AQR, Renaissance)의
  최신 Form 13F-HR 공시 접수일 및 포트폴리오 내 대상 종목 보유/옵션(Call/Put) 포지션 스캔
"""

import time
import xml.etree.ElementTree as ET
import requests
import pandas as pd
from config import TRACKED_FUNDS, SEC_USER_AGENT


def get_latest_13f_info_table_url(cik: str, headers: dict) -> tuple[str | None, str | None]:
    cik_padded = cik.zfill(10)
    sub_url = f"https://data.sec.gov/submissions/CIK{cik_padded}.json"
    resp = requests.get(sub_url, headers=headers, timeout=15)
    resp.raise_for_status()
    data = resp.json()

    recent = data.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    acc_nums = recent.get("accessionNumber", [])
    filing_dates = recent.get("filingDate", [])

    for form, acc, fdate in zip(forms, acc_nums, filing_dates):
        if form in ("13F-HR", "13F-HR/A"):
            acc_no_dash = acc.replace("-", "")
            cik_int = str(int(cik))
            idx_url = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_no_dash}/index.json"
            idx_resp = requests.get(idx_url, headers=headers, timeout=15)
            idx_resp.raise_for_status()
            items = idx_resp.json().get("directory", {}).get("item", [])
            for it in items:
                name = it.get("name", "")
                if name.endswith(".xml") and "primary_doc" not in name.lower():
                    xml_url = f"https://www.sec.gov/Archives/edgar/data/{cik_int}/{acc_no_dash}/{name}"
                    return xml_url, fdate
    return None, None


def fetch_hedge_fund_13f_summary(portfolio_items: list[dict]) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    반환값:
    1) 종목별 헤지펀드 보유 펀드 수 & 13F 콜옵션 보유 펀드 수 요약 DataFrame
    2) 펀드별 최신 13F 공시 접수 현황 DataFrame
    """
    headers = {"User-Agent": SEC_USER_AGENT, "Accept-Encoding": "gzip, deflate"}
    cusip_to_ticker = {item["cusip"]: item["ticker"] for item in portfolio_items}

    ticker_stats = {
        item["ticker"]: {"Hold_Funds": set(), "Call_Funds": set(), "Put_Funds": set(), "Total_Value_K": 0}
        for item in portfolio_items
    }
    filing_logs = []

    for fund_name, cik in TRACKED_FUNDS.items():
        try:
            xml_url, fdate = get_latest_13f_info_table_url(cik, headers)
            filing_logs.append({"Fund": fund_name, "CIK": cik, "Latest_13F_Date": fdate or "N/A", "Status": "수집 완료" if xml_url else "미확인"})
            if not xml_url:
                continue

            xml_resp = requests.get(xml_url, headers=headers, timeout=25)
            xml_resp.raise_for_status()
            root = ET.fromstring(xml_resp.content)

            for elem in root.iter():
                if elem.tag.endswith("infoTable"):
                    cusip_val = ""
                    val_usd_k = 0
                    put_call = ""
                    for child in elem:
                        tag = child.tag.split("}")[-1]
                        if tag == "cusip" and child.text:
                            cusip_val = child.text.strip()
                        elif tag == "value" and child.text:
                            val_usd_k = int(float(child.text.strip()))
                        elif tag == "putCall" and child.text:
                            put_call = child.text.strip().upper()

                    if cusip_val in cusip_to_ticker:
                        t = cusip_to_ticker[cusip_val]
                        ticker_stats[t]["Hold_Funds"].add(fund_name)
                        ticker_stats[t]["Total_Value_K"] += val_usd_k
                        if put_call == "CALL":
                            ticker_stats[t]["Call_Funds"].add(fund_name)
                        elif put_call == "PUT":
                            ticker_stats[t]["Put_Funds"].add(fund_name)
        except Exception as e:
            filing_logs.append({"Fund": fund_name, "CIK": cik, "Latest_13F_Date": "Error", "Status": str(e)[:40]})
        time.sleep(0.2)

    summary_rows = []
    for t, st in ticker_stats.items():
        summary_rows.append({
            "Ticker": t,
            "HF_Hold_Count": len(st["Hold_Funds"]),
            "HF_Call_Count": len(st["Call_Funds"]),
            "HF_Put_Count": len(st["Put_Funds"]),
            "HF_Hold_Names": ", ".join(sorted(st["Hold_Funds"])) if st["Hold_Funds"] else "-",
            "HF_13F_Value_M": round(st["Total_Value_K"] / 1000.0, 1),
        })

    return pd.DataFrame(summary_rows), pd.DataFrame(filing_logs)
