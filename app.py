# -*- coding: utf-8 -*-
"""
미래에셋증권 일산WM - 헤지펀드 13F · 옵션 매매동향 · 수익률 통합 트래커 (Streamlit)
"""

import os
import pandas as pd
import streamlit as st
from config import BRANCH_NAME, PORTFOLIO

OUTPUT_DIR = "output"

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

st.info("💡 주간 데이터와 통합 PPTX 파일은 매주 일요일 오전 9시(KST) GitHub Actions가 자동으로 수집해 저장소(output/)에 업데이트합니다.")

excel_path = os.path.join(OUTPUT_DIR, "ilsan_wm_weekly_history.xlsx")

if not os.path.exists(excel_path):
    st.warning("아직 `output/ilsan_wm_weekly_history.xlsx` 파일이 없습니다. GitHub Actions에서 워크플로우를 한 번 실행해 주세요.")
    st.stop()

xl_file = pd.ExcelFile(excel_path)
hist_df = pd.read_excel(xl_file, sheet_name="Weekly_History")
flow_df = pd.read_excel(xl_file, sheet_name="Latest_Flow_And_Returns")
filing_df = (
    pd.read_excel(xl_file, sheet_name="HedgeFund_13F_Logs")
    if "HedgeFund_13F_Logs" in xl_file.sheet_names
    else pd.DataFrame()
)
watch_df = (
    pd.read_excel(xl_file, sheet_name="Sector_25_Watchlist")
    if "Sector_25_Watchlist" in xl_file.sheet_names
    else pd.DataFrame()
)

latest = hist_df.iloc[-1]

c1, c2, c3, c4 = st.columns(4)
c1.metric("주간 수익률 (1W)", f"{latest['Port_1W']:+.2f}%")
c2.metric("월간 수익률 (1M)", f"{latest['Port_1M']:+.2f}%")
c3.metric(
    "누적 수익률 (2025~)",
    f"{latest['Port_Cum']:+.2f}%",
    f"SPY 대비 {latest['Alpha_vs_SPY_p']:+.2f}%p",
)
c4.metric(
    "1억 원 투자 현재 평가액",
    f"{latest['Eval_KRW_Man']:,.0f}만 원",
    f"기준일: {latest['Market_AsOf']}",
)


def build_cards(df_flow: pd.DataFrame, df_watch: pd.DataFrame, row_latest: pd.Series) -> list[tuple[str, str]]:
    cards = []
    if not df_flow.empty and "1W_Return" in df_flow.columns:
        non_etf = df_flow[~df_flow["Ticker"].isin(["SPY", "QQQ"])].copy()
        if non_etf.empty:
            non_etf = df_flow.copy()
        best_1w = non_etf.sort_values("1W_Return", ascending=False).iloc[0]
        worst_1w = non_etf.sort_values("1W_Return", ascending=True).iloc[0]
        cards.append((
            "01 · 주간 포트폴리오 수익률 및 주도주 점검",
            f"기준일({row_latest['Market_AsOf']}) 모델 포트폴리오는 주간 {row_latest['Port_1W']:+.2f}%, "
            f"2025년 이후 누적 {row_latest['Port_Cum']:+.2f}%(1억 원 기준 {row_latest['Eval_KRW_Man']:,.0f}만 원)를 기록했습니다. "
            f"주간 상대 강세 종목은 {best_1w['Ticker']}({best_1w['1W_Return']*100:+.2f}%), "
            f"단기 변동성 확대 종목은 {worst_1w['Ticker']}({worst_1w['1W_Return']*100:+.2f}%)입니다.",
        ))

        if "Call_Put_Vol_Ratio" in non_etf.columns:
            top_call = non_etf.sort_values("Call_Put_Vol_Ratio", ascending=False).head(3)
            call_txt = ", ".join([f"{r['Ticker']}({float(r['Call_Put_Vol_Ratio']):.2f}x)" for _, r in top_call.iterrows()])
            cards.append((
                "02 · 단기 옵션 매매동향 (콜/풋 거래량 비율)",
                f"근월물~차월물 옵션 체인 기준 콜/풋 거래량 비율 상위 종목은 {call_txt} 순입니다. "
                "콜 비율이 높고 거래량이 미결제약정을 웃도는 구간은 단기 상방 감마 수급 유입 여부를 함께 점검해야 합니다.",
            ))

        if "HF_Hold_Count" in non_etf.columns and "HF_Call_Count" in non_etf.columns:
            top_hf = non_etf.sort_values(["HF_Call_Count", "HF_Hold_Count"], ascending=False).head(3)
            hf_txt = ", ".join([
                f"{r['Ticker']}(보유 {int(r['HF_Hold_Count'])}곳·콜 {int(r['HF_Call_Count'])}곳)"
                for _, r in top_hf.iterrows()
            ])
            cards.append((
                "03 · 중기 대형 헤지펀드 SEC 13F 포지션 현황",
                f"주요 8대 멀티매니저·마켓메이커(Citadel, Millennium, Point72, Jane Street, SIG 등) 최신 13F 공시 기준 "
                f"기관 보유 및 콜옵션 포지션 상위 종목은 {hf_txt}입니다.",
            ))

    if not df_watch.empty and "Sector" in df_watch.columns and "1W" in df_watch.columns:
        sec_avg = df_watch.groupby("Sector", as_index=False)[["1W", "Cum"]].mean().sort_values("1W", ascending=False)
        top_sec = sec_avg.iloc[0]
        cards.append((
            "04 · 5대 주도 업종 로테이션 요약",
            f"5대 주도 업종 중 주간 평균 수익률 1위는 '{top_sec['Sector']}'({top_sec['1W']*100:+.2f}%, "
            f"누적 평균 {top_sec['Cum']*100:+.2f}%)입니다.",
        ))

    return cards


briefing_cards = build_cards(flow_df, watch_df, latest)
if briefing_cards:
    with st.expander("📝 이번 주 핵심 수급·수익률 자동 브리핑 요약 보기", expanded=True):
        for title_txt, body_txt in briefing_cards:
            st.markdown(
                f'<div class="insight-card"><strong>{title_txt}</strong><br>'
                f'<span style="font-size:0.95rem;">{body_txt}</span></div>',
                unsafe_allow_html=True,
            )

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 종목별 수익률 & 옵션·13F 수급",
    "🏦 주요 헤지펀드 13F 공시 현황",
    "📈 주차별 누적 수익률 추이",
    "⚙️ 포트폴리오 비중 편집기",
])

with tab1:
    wanted_cols = [
        "Ticker", "Sector", "Weight", "1W_Return", "1M_Return", "Cum_Return",
        "Call_Put_Vol_Ratio", "Unusual_Call_Strikes", "Option_Signal",
        "HF_Hold_Count", "HF_Call_Count", "HF_Hold_Names",
    ]
    avail_cols = [c for c in wanted_cols if c in flow_df.columns]
    disp = flow_df[avail_cols].copy()
    for col in ["Weight", "1W_Return", "1M_Return", "Cum_Return"]:
        if col in disp.columns:
            disp[col] = disp[col].map(
                lambda x: f"{float(x)*100:+.2f}%" if col != "Weight" else f"{float(x)*100:.0f}%"
            )
    st.dataframe(disp, use_container_width=True, hide_index=True)

with tab2:
    if not filing_df.empty:
        st.dataframe(filing_df, use_container_width=True, hide_index=True)
    else:
        st.info("13F 공시 로그 시트가 아직 없습니다.")

with tab3:
    chart_cols = [c for c in ["Port_Cum", "SPY_Cum"] if c in hist_df.columns]
    if chart_cols and "Market_AsOf" in hist_df.columns:
        st.line_chart(hist_df.set_index("Market_AsOf")[chart_cols])

with tab4:
    st.caption("비중(%)을 조정해 합계 100%를 맞춘 뒤, 아래 생성된 PORTFOLIO 코드를 복사해 config.py에 반영할 수 있습니다.")
    edit_df = pd.DataFrame(PORTFOLIO).copy()
    edit_df["weight_pct"] = edit_df["weight"] * 100.0
    cols_show = [c for c in ["ticker", "name", "sector", "weight_pct", "eps_2027", "cusip"] if c in edit_df.columns]
    edited = st.data_editor(edit_df[cols_show], use_container_width=True, hide_index=True)
    total_w = float(edited["weight_pct"].sum())
    st.write(
        f"**현재 입력된 비중 합계:** `{total_w:.1f}%` "
        + ("✅ (100% 완료)" if abs(total_w - 100.0) < 0.01 else "⚠️ 100%에 맞춰 주세요")
    )

    lines = ["PORTFOLIO = ["]
    for _, r in edited.iterrows():
        lines.append(
            f'    {{"ticker": "{r.get("ticker", "")}", "cusip": "{r.get("cusip", "")}", '
            f'"name": "{r.get("name", "")}", "sector": "{r.get("sector", "")}", '
            f'"weight": {round(float(r.get("weight_pct", 0))/100.0, 4)}, "eps_2027": "{r.get("eps_2027", "")}"}},'
        )
    lines.append("]")
    st.code("\n".join(lines), language="python")

if os.path.exists(OUTPUT_DIR):
    pptx_files = sorted([f for f in os.listdir(OUTPUT_DIR) if f.endswith(".pptx")])
    if pptx_files:
        latest_pptx = os.path.join(OUTPUT_DIR, pptx_files[-1])
        with open(latest_pptx, "rb") as f:
            st.download_button(
                label=f"📥 최신 일산WM 주간 PPTX 다운로드 ({pptx_files[-1]})",
                data=f,
                file_name=pptx_files[-1],
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )
