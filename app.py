# -*- coding: utf-8 -*-
"""
미래에셋증권 일산WM - 헤지펀드 13F · 옵션 매매동향 · 수익률 통합 트래커 (Streamlit)
실행 명령: streamlit run app.py
"""

import os
import pandas as pd
import streamlit as st
from config import BRANCH_NAME
from update_weekly import main as run_weekly_update, OUTPUT_DIR

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
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    f'<div class="header-banner"><h2>🍊 {BRANCH_NAME} | 헤지펀드(13F) · 옵션 수급 · 주간/월간/누적 수익률 트래커</h2>'
    f'<p style="margin:0;">매주 일요일 GitHub Actions 자동 업데이트 · 주가 + 옵션 콜/풋 플로우 + SEC 13F 기관 포지션 통합 대시보드</p></div>',
    unsafe_allow_html=True,
)

if st.button("🔄 지금 즉시 주가 · 옵션 수급 · 13F 공시 & PPTX 업데이트", type="primary"):
    with st.spinner("미국 증시 종가, 옵션 체인, SEC 13F 공시 데이터를 수집 중입니다..."):
        run_weekly_update()
    st.success("최신 수급/수익률 데이터와 PPTX 리포트가 업데이트되었습니다!")

excel_path = os.path.join(OUTPUT_DIR, "ilsan_wm_weekly_history.xlsx")
if not os.path.exists(excel_path):
    run_weekly_update()

hist_df = pd.read_excel(excel_path, sheet_name="Weekly_History")
flow_df = pd.read_excel(excel_path, sheet_name="Latest_Flow_And_Returns")
filing_df = pd.read_excel(excel_path, sheet_name="HedgeFund_13F_Logs")
latest = hist_df.iloc[-1]

c1, c2, c3, c4 = st.columns(4)
c1.metric("주간 수익률 (1W)", f"{latest['Port_1W']:+.2f}%")
c2.metric("월간 수익률 (1M)", f"{latest['Port_1M']:+.2f}%")
c3.metric("누적 수익률 (2025~)", f"{latest['Port_Cum']:+.2f}%", f"SPY 대비 {latest['Alpha_vs_SPY_p']:+.2f}%p")
c4.metric("1억 원 투자 현재 평가액", f"{latest['Eval_KRW_Man']:,.0f}만 원", f"기준일: {latest['Market_AsOf']}")

tab1, tab2, tab3 = st.tabs(["📊 종목별 수익률 & 옵션·13F 수급", "🏦 주요 헤지펀드 13F 공시 현황", "📈 주차별 누적 수익률 추이"])

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

pptx_files = sorted([f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pptx")])
if pptx_files:
    latest_pptx = os.path.join(OUTPUT_DIR, pptx_files[-1])
    with open(latest_pptx, "rb") as f:
        st.download_button(
            label=f"📥 최신 일산WM 주간 수급·수익률 PPTX 다운로드 ({pptx_files[-1]})",
            data=f,
            file_name=pptx_files[-1],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
