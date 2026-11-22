# -*- coding: utf-8 -*-
"""
미래에셋증권 일산WM - 헤지펀드 13F · 옵션 매매동향 · 수익률 통합 트래커 v2.0 (Streamlit)
"""

import os
import pandas as pd
import streamlit as st
from config import BRANCH_NAME, PORTFOLIO
from update_weekly import main as run_weekly_update, build_weekly_commentary, OUTPUT_DIR

st.set_page_config(page_title=f"{BRANCH_NAME} 수급·수익률 트래커", layout="wide")

st.markdown(
    """
    <style>
    .stApp { background-color: #FFF9F5; }
    .header-banner {
        background: linear-gradient(90deg, #F58220 0%, #D96B0C 100%);
        padding: 20px 28px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
    }
    .insight-card {
        background-color: white;
        border-left: 5px solid #F58220;
        padding: 14px 18px;
        border-radius: 8px;
        margin-bottom: 10px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.06);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f'<div class="header-banner"><h2>🍊 {BRANCH_NAME} | 헤지펀드(13F) · 옵션 수급 · 주간/월간/누적 수익률 트래커</h2>'
    f'<p style="margin:0;">매주 일요일 자동 업데이트 · 주간 핵심 인사이트 브리핑 + 옵션·13F 수급 + 통합 PPTX 원클릭 다운로드</p></div>',
    unsafe_allow_html=True,
)

if st.button("🔄 지금 즉시 주가 · 옵션 수급 · 13F 공시 & 통합 PPTX 업데이트", type="primary"):
    with st.spinner("미국 증시 종가, 옵션 체인, SEC 13F 공시 데이터를 수집 중입니다..."):
        run_weekly_update()
    st.success("최신 주간 요약 브리핑과 통합 PPTX 리포트(1~18p)가 업데이트되었습니다!")

excel_path = os.path.join(OUTPUT_DIR, "ilsan_wm_weekly_history.xlsx")
if not os.path.exists(excel_path):
    run_weekly_update()

hist_df = pd.read_excel(excel_path, sheet_name="Weekly_History")
flow_df = pd.read_excel(excel_path, sheet_name="Latest_Flow_And_Returns")
filing_df = pd.read_excel(excel_path, sheet_name="HedgeFund_13F_Logs")
watch_df = pd.read_excel(excel_path, sheet_name="Sector_25_Watchlist") if "Sector_25_Watchlist" in pd.ExcelFile(excel_path).sheet_names else pd.DataFrame()
latest = hist_df.iloc[-1]

c1, c2, c3, c4 = st.columns(4)
c1.metric("주간 수익률 (1W)", f"{latest['Port_1W']:+.2f}%")
c2.metric("월간 수익률 (1M)", f"{latest['Port_1M']:+.2f}%")
c3.metric("누적 수익률 (2025~)", f"{latest['Port_Cum']:+.2f}%", f"SPY 대비 {latest['Alpha_vs_SPY_p']:+.2f}%p")
c4.metric("1억 원 투자 현재 평가액", f"{latest['Eval_KRW_Man']:,.0f}만 원", f"기준일: {latest['Market_AsOf']}")

summary_dict = {
    "as_of": latest["Market_AsOf"],
    "port_1w": latest["Port_1W"] / 100.0,
    "port_1m": latest["Port_1M"] / 100.0,
    "port_cum": latest["Port_Cum"] / 100.0,
    "eval_krw": latest["Eval_KRW_Man"] * 10000.0,
    "spy_1w": flow_df.loc[flow_df["Ticker"] == "SPY", "1W_Return"].iloc[0],
    "alpha_vs_spy": latest["Alpha_vs_SPY_p"] / 100.0,
}

with st.expander("📝 이번 주 핵심 수급·수익률 자동 브리핑 요약 보기", expanded=True):
    for title_txt, body_txt in build_weekly_commentary(flow_df, watch_df, summary_dict):
        st.markdown(f'<div class="insight-card"><strong>{title_txt}</strong><br><span style="font-size:0.95rem;">{body_txt}</span></div>', unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 종목별 수익률 & 옵션·13F 수급",
    "🏦 주요 헤지펀드 13F 공시 현황",
    "📈 주차별 누적 수익률 추이",
    "⚙️ 포트폴리오 비중 편집기",
])

with tab1:
    disp = flow_df[[
        "Ticker", "Sector", "Weight", "1W_Return", "1M_Return", "Cum_Return",
        "Call_Put_Vol_Ratio", "Unusual_Call_Strikes", "Option_Signal",
        "HF_Hold_Count", "HF_Call_Count", "HF_Hold_Names"
    ]].copy()
    for col in ["Weight", "1W_Return", "1M_Return", "Cum_Return"]:
        disp[col] = disp[col].map(lambda x: f"{x*100:+.2f}%" if col != "Weight" else f"{x*100:.0f}%")
    st.dataframe(disp, use_container_width=True, hide_index=True)

with tab2:
    st.dataframe(filing_df, use_container_width=True, hide_index=True)

with tab3:
    st.line_chart(hist_df.set_index("Market_AsOf")[["Port_Cum", "SPY_Cum"]])

with tab4:
    st.caption("비중(%)을 조정해 비중 합계 100%를 맞춘 뒤, 아래 생성된 config.py 코드를 복사해 GitHub에 붙여넣으면 다음 주부터 새 비중으로 자동 집계됩니다.")
    edit_df = pd.DataFrame(PORTFOLIO)[["ticker", "cusip", "name", "sector", "weight", "eps_2027"]].copy()
    edit_df["weight_pct"] = edit_df["weight"] * 100.0
    edited = st.data_editor(
        edit_df[["ticker", "name", "sector", "weight_pct", "eps_2027", "cusip"]],
        use_container_width=True,
        hide_index=True,
    )
    total_w = edited["weight_pct"].sum()
    st.write(f"**현재 입력된 비중 합계:** `{total_w:.1f}%` {'✅ (정확히 100%입니다)' if abs(total_w - 100.0) < 0.01 else '⚠️ 합계가 100%가 되도록 조정해 주세요'}")

    lines = ["PORTFOLIO = ["]
    for _, r in edited.iterrows():
        lines.append(
            f'    {{"ticker": "{r["ticker"]}", "cusip": "{r["cusip"]}", "name": "{r["name"]}", '
            f'"sector": "{r["sector"]}", "weight": {round(r["weight_pct"]/100.0, 4)}, "eps_2027": "{r["eps_2027"]}"}},'
        )
    lines.append("]")
    st.code("\n".join(lines), language="python")

pptx_files = sorted([f for f in os.listdir(OUTPUT_DIR) if f.startswith("mirae_ilsan_wm_integrated_report_") and f.endswith(".pptx")])
if pptx_files:
    latest_pptx = os.path.join(OUTPUT_DIR, pptx_files[-1])
    with open(latest_pptx, "rb") as f:
        st.download_button(
            label=f"📥 최신 일산WM 통합 PPTX 리포트 다운로드 (1~18p · {pptx_files[-1]})",
            data=f,
            file_name=pptx_files[-1],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
