# -*- coding: utf-8 -*-
"""
매주 일요일 자동 실행 통합 엔진 (update_weekly.py)
1) 주가 수익률: 주간(1W) · 월간(1M) · 연초대비(YTD) · 2025년 초 이후 전체 누적 수익률
2) 옵션 매매동향: 콜/풋 거래량 비율, 콜/풋 미결제약정(OI) 비율, 이상 콜옵션 수급 시그널
3) 헤지펀드 동향: SEC EDGAR 13F 공식 공시 기반 대형 8대 헤지펀드 보유 및 13F 콜옵션 포지션 집계
4) 산출물: Excel 누적 로그(ilsan_wm_weekly_history.xlsx) + 미래에셋증권 일산WM 오렌지 테마 PPTX 자동 생성
"""

import os
from datetime import datetime, timedelta
import pandas as pd
import yfinance as yf
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from config import (
    PORTFOLIO,
    START_DATE,
    INITIAL_CAPITAL_KRW,
    BRAND_PRIMARY_HEX,
    BRAND_NAVY_HEX,
    BRANCH_NAME,
)
from options_flow import fetch_options_metrics
from sec_13f import fetch_hedge_fund_13f_summary

OUTPUT_DIR = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)


def hex_to_rgb(hex_str: str) -> RGBColor:
    return RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))


def fetch_price_history(tickers: list[str], start_date: str) -> pd.DataFrame:
    end_dt = (datetime.now() + timedelta(days=1)).strftime("%Y-%m-%d")
    raw = yf.download(tickers, start=start_date, end=end_dt, auto_adjust=True, progress=False)
    if isinstance(raw.columns, pd.MultiIndex):
        prices = raw["Close"].copy()
    else:
        prices = raw[["Close"]].copy()
        prices.columns = tickers
    return prices.ffill().dropna(how="all")


def compute_period_returns(prices: pd.DataFrame) -> pd.DataFrame:
    latest_date = prices.index[-1]
    ytd_prices = prices.loc[prices.index >= f"{latest_date.year}-01-01"]

    rows = []
    for item in PORTFOLIO:
        t = item["ticker"]
        s = prices[t].dropna()
        p_now = float(s.iloc[-1])
        p_start = float(s.iloc[0])
        p_1w = float(s.iloc[-6]) if len(s) >= 6 else p_start
        p_1m = float(s.iloc[-22]) if len(s) >= 22 else p_start
        s_ytd = ytd_prices[t].dropna()
        p_ytd_base = float(s_ytd.iloc[0]) if len(s_ytd) > 0 else p_start

        rows.append({
            "Ticker": t,
            "Name": item["name"],
            "Sector": item["sector"],
            "Weight": item["weight"],
            "Price": round(p_now, 2),
            "1W_Return": (p_now / p_1w) - 1.0,
            "1M_Return": (p_now / p_1m) - 1.0,
            "YTD_Return": (p_now / p_ytd_base) - 1.0,
            "Cum_Return": (p_now / p_start) - 1.0,
            "EPS_2027": item["eps_2027"],
            "AsOf": latest_date.strftime("%Y-%m-%d"),
        })
    return pd.DataFrame(rows)


def compute_portfolio_summary(df: pd.DataFrame) -> dict:
    w = df["Weight"]
    cur = w * (1.0 + df["Cum_Return"])
    tot_cur = cur.sum()

    v_1w = cur / (1.0 + df["1W_Return"])
    v_1m = cur / (1.0 + df["1M_Return"])
    v_ytd = cur / (1.0 + df["YTD_Return"])

    port_1w = (tot_cur / v_1w.sum()) - 1.0
    port_1m = (tot_cur / v_1m.sum()) - 1.0
    port_ytd = (tot_cur / v_ytd.sum()) - 1.0
    port_cum = tot_cur - 1.0

    spy = df[df["Ticker"] == "SPY"].iloc[0]
    return {
        "as_of": df["AsOf"].iloc[0],
        "port_1w": port_1w,
        "port_1m": port_1m,
        "port_ytd": port_ytd,
        "port_cum": port_cum,
        "eval_krw": INITIAL_CAPITAL_KRW * (1.0 + port_cum),
        "spy_1w": spy["1W_Return"],
        "spy_1m": spy["1M_Return"],
        "spy_cum": spy["Cum_Return"],
        "alpha_vs_spy": port_cum - spy["Cum_Return"],
    }


def save_excel_history(merged_df: pd.DataFrame, filing_df: pd.DataFrame, summary: dict) -> str:
    excel_path = os.path.join(OUTPUT_DIR, "ilsan_wm_weekly_history.xlsx")
    summary_row = pd.DataFrame([{
        "Update_Date": datetime.now().strftime("%Y-%m-%d"),
        "Market_AsOf": summary["as_of"],
        "Port_1W": round(summary["port_1w"] * 100, 2),
        "Port_1M": round(summary["port_1m"] * 100, 2),
        "Port_YTD": round(summary["port_ytd"] * 100, 2),
        "Port_Cum": round(summary["port_cum"] * 100, 2),
        "Eval_KRW_Man": round(summary["eval_krw"] / 10000, 0),
        "SPY_Cum": round(summary["spy_cum"] * 100, 2),
        "Alpha_vs_SPY_p": round(summary["alpha_vs_spy"] * 100, 2),
    }])

    if os.path.exists(excel_path):
        old_hist = pd.read_excel(excel_path, sheet_name="Weekly_History")
        hist_df = pd.concat([old_hist, summary_row], ignore_index=True).drop_duplicates(subset=["Market_AsOf"], keep="last")
    else:
        hist_df = summary_row

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        hist_df.to_excel(writer, sheet_name="Weekly_History", index=False)
        merged_df.to_excel(writer, sheet_name="Latest_Flow_And_Returns", index=False)
        filing_df.to_excel(writer, sheet_name="HedgeFund_13F_Logs", index=False)

    return excel_path


def generate_weekly_pptx(merged_df: pd.DataFrame, summary: dict) -> str:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    orange = hex_to_rgb(BRAND_PRIMARY_HEX)
    navy = hex_to_rgb(BRAND_NAVY_HEX)
    white = RGBColor(255, 255, 255)
    dark_text = RGBColor(30, 35, 45)
    light_bg = RGBColor(255, 248, 242)

    # Slide 1: Cover
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    s1.background.fill.solid()
    s1.background.fill.fore_color.rgb = orange

    tb = s1.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(11.5), Inches(3.8))
    tf = tb.text_frame
    tf.word_wrap = True
    p0 = tf.paragraphs[0]
    p0.text = f"{BRANCH_NAME} | WEEKLY HEDGE FUND & OPTIONS FLOW REPORT"
    p0.font.size = Pt(16)
    p0.font.bold = True
    p0.font.color.rgb = navy

    p1 = tf.add_paragraph()
    p1.text = "미국 주도 세력(헤지펀드 13F · 콜옵션 수급) 주간 트래커"
    p1.font.size = Pt(34)
    p1.font.bold = True
    p1.font.color.rgb = white
    p1.space_before = Pt(14)

    p2 = tf.add_paragraph()
    p2.text = f"기준일: {summary['as_of']} · 주간(1W)/월간(1M)/누적 수익률 + 옵션/13F 수급 통합본"
    p2.font.size = Pt(18)
    p2.font.color.rgb = white
    p2.space_before = Pt(18)

    # Slide 2: Return & Flow Integrated Table
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    s2.background.fill.solid()
    s2.background.fill.fore_color.rgb = light_bg

    hdr = s2.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12.1), Inches(0.65))
    hp = hdr.text_frame.paragraphs[0]
    hp.text = f"[{BRANCH_NAME}] 주간(1W)·월간(1M)·누적 수익률 및 헤지펀드·옵션 매매동향 ({summary['as_of']})"
    hp.font.size = Pt(21)
    hp.font.bold = True
    hp.font.color.rgb = navy

    rows_cnt = len(merged_df) + 1
    tbl_shape = s2.shapes.add_table(rows_cnt, 9, Inches(0.6), Inches(1.2), Inches(12.13), Inches(5.7))
    tbl = tbl_shape.table
    headers = ["티커", "비중", "주간(1W)", "월간(1M)", "누적(2025~)", "콜/풋 비율", "옵션 수급 시그널", "주요 HF 보유", "13F 콜 보유"]
    for c_idx, h_text in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = navy
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.size = Pt(10.5)
        p.font.bold = True
        p.font.color.rgb = white
        p.alignment = PP_ALIGN.CENTER

    for r_idx, row in merged_df.iterrows():
        vals = [
            row["Ticker"],
            f"{row['Weight']*100:.0f}%",
            f"{row['1W_Return']*100:+.2f}%",
            f"{row['1M_Return']*100:+.2f}%",
            f"{row['Cum_Return']*100:+.2f}%",
            f"{row['Call_Put_Vol_Ratio']:.2f}x",
            row["Option_Signal"],
            f"{row['HF_Hold_Count']}곳",
            f"{row['HF_Call_Count']}곳",
        ]
        for c_idx, v in enumerate(vals):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = white
            p = cell.text_frame.paragraphs[0]
            p.text = str(v)
            p.font.size = Pt(10)
            p.font.color.rgb = dark_text
            p.alignment = PP_ALIGN.CENTER

    out_pptx = os.path.join(OUTPUT_DIR, f"mirae_ilsan_wm_flow_report_{summary['as_of']}.pptx")
    prs.save(out_pptx)
    return out_pptx


def main():
    tickers = [x["ticker"] for x in PORTFOLIO]
    print("1/4 주가 시계열 및 1W/1M/누적 수익률 계산 중...")
    prices = fetch_price_history(tickers, START_DATE)
    ret_df = compute_period_returns(prices)
    summary = compute_portfolio_summary(ret_df)

    print("2/4 옵션 체인(콜/풋 비율 · 이상 콜거래) 수집 중...")
    opt_df = fetch_options_metrics(tickers)

    print("3/4 SEC EDGAR 13F 헤지펀드 보유 및 콜옵션 공시 스캔 중...")
    hf_df, filing_df = fetch_hedge_fund_13f_summary(PORTFOLIO)

    merged_df = ret_df.merge(opt_df, on="Ticker", how="left").merge(hf_df, on="Ticker", how="left")

    print("4/4 Excel 누적 히스토리 및 PPTX 리포트 저장 중...")
    excel_file = save_excel_history(merged_df, filing_df, summary)
    pptx_file = generate_weekly_pptx(merged_df, summary)
    print(f"[완료] Excel: {excel_file}")
    print(f"[완료] PPTX : {pptx_file}")


if __name__ == "__main__":
    main()
