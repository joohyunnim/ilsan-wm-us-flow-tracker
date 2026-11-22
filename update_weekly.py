# -*- coding: utf-8 -*-
"""
미래에셋증권 일산WM - 마스터 전략 덱(base_deck.pptx) + 주간 자동 업데이트 리포트 병합 엔진
- 앞부분(1~12p): 고정 마스터 전략 슬라이드(base_deck.pptx) 유지
- 뒷부분(13p~): 매주 일요일 자동 수집된 주간(1W)·월간(1M)·누적 수익률 + 옵션 수급 + SEC 13F + 25종목 업종 스캔 슬라이드 자동 부착
"""

import os
import copy
from datetime import datetime, timedelta
import pandas as pd
import yfinance as yf
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from config import (
    PORTFOLIO,
    WATCHLIST_5_SECTORS,
    START_DATE,
    INITIAL_CAPITAL_KRW,
    BRAND_PRIMARY_HEX,
    BRAND_NAVY_HEX,
    BRANCH_NAME,
)
from options_flow import fetch_options_metrics
from sec_13f import fetch_hedge_fund_13f_summary

OUTPUT_DIR = "output"
BASE_DECK_PATH = "base_deck.pptx"
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


def compute_watchlist_returns(prices: pd.DataFrame) -> pd.DataFrame:
    latest_date = prices.index[-1]
    rows = []
    for sector_name, tickers in WATCHLIST_5_SECTORS.items():
        for t in tickers:
            if t not in prices.columns:
                continue
            s = prices[t].dropna()
            if len(s) < 2:
                continue
            p_now = float(s.iloc[-1])
            p_start = float(s.iloc[0])
            p_1w = float(s.iloc[-6]) if len(s) >= 6 else p_start
            p_1m = float(s.iloc[-22]) if len(s) >= 22 else p_start
            rows.append({
                "Sector": sector_name,
                "Ticker": t,
                "Price": round(p_now, 2),
                "1W": (p_now / p_1w) - 1.0,
                "1M": (p_now / p_1m) - 1.0,
                "Cum": (p_now / p_start) - 1.0,
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


def save_excel_history(merged_df: pd.DataFrame, watch_df: pd.DataFrame, filing_df: pd.DataFrame, summary: dict) -> str:
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
        watch_df.to_excel(writer, sheet_name="Sector_25_Watchlist", index=False)
        filing_df.to_excel(writer, sheet_name="HedgeFund_13F_Logs", index=False)

    return excel_path


def delete_last_slide(prs: Presentation):
    """마스터 덱의 맨 마지막 '감사합니다' 클로징 슬라이드를 잠시 제거해 중간에 주간 리포트를 삽입할 수 있게 함"""
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    if len(slides) > 0:
        rId = slides[-1].rId
        prs.part.drop_rel(rId)
        xml_slides.remove(slides[-1])


def generate_merged_pptx(merged_df: pd.DataFrame, watch_df: pd.DataFrame, summary: dict) -> str:
    orange = hex_to_rgb(BRAND_PRIMARY_HEX)
    navy = hex_to_rgb(BRAND_NAVY_HEX)
    white = RGBColor(255, 255, 255)
    dark_text = RGBColor(30, 35, 45)
    light_bg = RGBColor(255, 248, 242)
    zebra_bg = RGBColor(255, 243, 232)

    # 1. 마스터 전략 덱(base_deck.pptx)이 있으면 불러오고 마지막 '감사합니다' 슬라이드 제거 후 뒤에 이어붙임
    if os.path.exists(BASE_DECK_PATH):
        prs = Presentation(BASE_DECK_PATH)
        if len(prs.slides) >= 13:
            delete_last_slide(prs)
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    # [추가 슬라이드 1] PART 05 간지 (주간 자동 업데이트 섹션 시작)
    sec_slide = prs.slides.add_slide(blank_layout)
    sec_slide.background.fill.solid()
    sec_slide.background.fill.fore_color.rgb = orange

    tb_sec = sec_slide.shapes.add_textbox(Inches(0.9), Inches(2.2), Inches(11.5), Inches(3.2))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    p_tag = tf_sec.paragraphs[0]
    p_tag.text = f"PART 05 · WEEKLY LIVE TRACKER ({summary['as_of']} 기준 자동 갱신)"
    p_tag.font.size = Pt(16)
    p_tag.font.bold = True
    p_tag.font.color.rgb = navy

    p_main = tf_sec.add_paragraph()
    p_main.text = "주간(1W) · 월간(1M) · 누적 수익률 및\n헤지펀드 13F · 콜옵션 실시간 수급 현황"
    p_main.font.size = Pt(34)
    p_main.font.bold = True
    p_main.font.color.rgb = white
    p_main.space_before = Pt(14)

    # [추가 슬라이드 2] 포트폴리오 주간·월간·누적 수익률 & 옵션·13F 통합 대시보드
    s_port = prs.slides.add_slide(blank_layout)
    s_port.background.fill.solid()
    s_port.background.fill.fore_color.rgb = light_bg

    hdr = s_port.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.55))
    hp = hdr.text_frame.paragraphs[0]
    hp.text = f"05 · 모델 포트폴리오 주간·월간·누적 성과 및 옵션·13F 수급 ({summary['as_of']})"
    hp.font.size = Pt(20)
    hp.font.bold = True
    hp.font.color.rgb = navy

    kpis = [
        ("주간 수익률 (1W)", f"{summary['port_1w']*100:+.2f}%", f"SPY {summary['spy_1w']*100:+.2f}%"),
        ("월간 수익률 (1M)", f"{summary['port_1m']*100:+.2f}%", f"SPY {summary['spy_1m']*100:+.2f}%"),
        ("전체 누적 (2025~)", f"{summary['port_cum']*100:+.2f}%", f"SPY 대비 {summary['alpha_vs_spy']*100:+.2f}%p"),
        ("1억 투자 현재 평가액", f"{summary['eval_krw']/10000:,.0f}만 원", f"기준일 {summary['as_of']}"),
    ]
    for idx, (label, val, sub) in enumerate(kpis):
        x = Inches(0.6 + idx * 3.1)
        card = s_port.shapes.add_shape(1, x, Inches(1.0), Inches(2.9), Inches(1.15))
        card.fill.solid()
        card.fill.fore_color.rgb = white
        card.line.color.rgb = orange
        ctf = card.text_frame
        ctf.word_wrap = True
        cp0 = ctf.paragraphs[0]
        cp0.text = label
        cp0.font.size = Pt(11)
        cp0.font.bold = True
        cp0.font.color.rgb = navy
        cp1 = ctf.add_paragraph()
        cp1.text = val
        cp1.font.size = Pt(20)
        cp1.font.bold = True
        cp1.font.color.rgb = orange
        cp2 = ctf.add_paragraph()
        cp2.text = sub
        cp2.font.size = Pt(10)
        cp2.font.color.rgb = dark_text

    rows_cnt = len(merged_df) + 1
    tbl_shape = s_port.shapes.add_table(rows_cnt, 9, Inches(0.6), Inches(2.3), Inches(12.13), Inches(4.8))
    tbl = tbl_shape.table
    headers = ["티커", "비중", "주간(1W)", "월간(1M)", "누적(2025~)", "콜/풋 비율", "옵션 수급 시그널", "주요 HF 보유", "13F 콜 보유"]
    for c_idx, h_text in enumerate(headers):
        cell = tbl.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = navy
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.size = Pt(10)
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
        bg_c = white if r_idx % 2 == 0 else zebra_bg
        for c_idx, v in enumerate(vals):
            cell = tbl.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_c
            p = cell.text_frame.paragraphs[0]
            p.text = str(v)
            p.font.size = Pt(9.5)
            p.font.color.rgb = dark_text
            p.alignment = PP_ALIGN.CENTER

    # [추가 슬라이드 3] 5대 주도 업종 25종목 주간·월간·누적 수익률 & 옵션 스캔 요약
    if not watch_df.empty:
        s_watch = prs.slides.add_slide(blank_layout)
        s_watch.background.fill.solid()
        s_watch.background.fill.fore_color.rgb = light_bg

        whdr = s_watch.shapes.add_textbox(Inches(0.6), Inches(0.35), Inches(12.1), Inches(0.55))
        whp = whdr.text_frame.paragraphs[0]
        whp.text = f"05 · 5대 주도 업종 핵심 25종목 주간·월간·누적 모멘텀 스캔 ({summary['as_of']})"
        whp.font.size = Pt(20)
        whp.font.bold = True
        whp.font.color.rgb = navy

        # 상위 모멘텀 15종목 테이블 표시 (가독성 최적화)
        top_watch = watch_df.sort_values(by="1W", ascending=False).head(15).reset_index(drop=True)
        w_tbl_shape = s_watch.shapes.add_table(len(top_watch) + 1, 6, Inches(0.6), Inches(1.05), Inches(12.13), Inches(5.95))
        wtbl = w_tbl_shape.table
        wheaders = ["소속 주도 업종", "티커", "현재가($)", "주간(1W)", "월간(1M)", "누적(2025~)"]
        for c_idx, h_text in enumerate(wheaders):
            cell = wtbl.cell(0, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = navy
            p = cell.text_frame.paragraphs[0]
            p.text = h_text
            p.font.size = Pt(10)
            p.font.bold = True
            p.font.color.rgb = white
            p.alignment = PP_ALIGN.CENTER

        for r_idx, row in top_watch.iterrows():
            vals = [
                row["Sector"],
                row["Ticker"],
                f"${row['Price']:,.2f}",
                f"{row['1W']*100:+.2f}%",
                f"{row['1M']*100:+.2f}%",
                f"{row['Cum']*100:+.2f}%",
            ]
            bg_c = white if r_idx % 2 == 0 else zebra_bg
            for c_idx, v in enumerate(vals):
                cell = wtbl.cell(r_idx + 1, c_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = bg_c
                p = cell.text_frame.paragraphs[0]
                p.text = str(v)
                p.font.size = Pt(9.5)
                p.font.color.rgb = dark_text
                p.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.CENTER

    # [마지막 슬라이드] 미래에셋증권 일산WM 통합 클로징 슬라이드
    s_end = prs.slides.add_slide(blank_layout)
    s_end.background.fill.solid()
    s_end.background.fill.fore_color.rgb = orange

    etb = s_end.shapes.add_textbox(Inches(1.0), Inches(2.5), Inches(11.3), Inches(2.5))
    etf = etb.text_frame
    etf.word_wrap = True
    ep0 = etf.paragraphs[0]
    ep0.text = "감사합니다"
    ep0.font.size = Pt(40)
    ep0.font.bold = True
    ep0.font.color.rgb = white

    ep1 = etf.add_paragraph()
    ep1.text = f"{BRANCH_NAME} · 기준일 {summary['as_of']} 자동 업데이트 완료"
    ep1.font.size = Pt(18)
    ep1.font.bold = True
    ep1.font.color.rgb = navy
    ep1.space_before = Pt(12)

    out_pptx = os.path.join(OUTPUT_DIR, f"mirae_ilsan_wm_integrated_report_{summary['as_of']}.pptx")
    prs.save(out_pptx)
    return out_pptx


def main():
    port_tickers = [x["ticker"] for x in PORTFOLIO]
    watch_tickers = list({t for t_list in WATCHLIST_5_SECTORS.values() for t in t_list})
    all_tickers = sorted(list(set(port_tickers + watch_tickers)))

    print("1/4 포트폴리오 및 5대 업종 25종목 주가 시계열 수집 중...")
    prices = fetch_price_history(all_tickers, START_DATE)
    ret_df = compute_period_returns(prices)
    watch_df = compute_watchlist_returns(prices)
    summary = compute_portfolio_summary(ret_df)

    print("2/4 옵션 체인(콜/풋 비율 · 이상 콜거래) 수집 중...")
    opt_df = fetch_options_metrics(port_tickers)

    print("3/4 SEC EDGAR 13F 헤지펀드 보유 및 콜옵션 공시 스캔 중...")
    hf_df, filing_df = fetch_hedge_fund_13f_summary(PORTFOLIO)

    merged_df = ret_df.merge(opt_df, on="Ticker", how="left").merge(hf_df, on="Ticker", how="left")

    print("4/4 Excel 누적 히스토리 및 통합 PPTX(앞 전략장표 + 뒤 주간업데이트) 저장 중...")
    excel_file = save_excel_history(merged_df, watch_df, filing_df, summary)
    pptx_file = generate_merged_pptx(merged_df, watch_df, summary)
    print(f"[완료] Excel: {excel_file}")
    print(f"[완료] 통합 PPTX: {pptx_file}")


if __name__ == "__main__":
    main()
