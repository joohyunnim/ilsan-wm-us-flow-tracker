# -*- coding: utf-8 -*-
"""
미래에셋증권 일산WM - 미국 시장 주도 세력(헤지펀드 13F · 옵션 플로우 · 주가) 주간 자동화 설정
"""

BRANCH_NAME = "미래에셋증권 일산WM"
BRAND_PRIMARY_HEX = "F58220"  # 미래에셋 오렌지
BRAND_NAVY_HEX = "002C5F"     # 미래에셋 딥 네이비
START_DATE = "2025-01-02"
INITIAL_CAPITAL_KRW = 100_000_000  # 1억 원 기준

# SEC EDGAR API 요청 시 필수 User-Agent (본인 이메일로 수정 권장)
SEC_USER_AGENT = "MiraeAssetIlsanWM/1.0 (ilsan_wm_quant@example.com)"

# 추적 대상 주요 대형 헤지펀드 및 마켓메이커 CIK 번호 (SEC 13F 공시 기준)
TRACKED_FUNDS = {
    "Citadel Advisors": "0001423053",
    "Millennium Management": "0001273087",
    "Point72 Asset Management": "0001603466",
    "Jane Street Group": "0001595888",
    "Susquehanna (SIG)": "0001446194",
    "Bridgewater Associates": "0001350694",
    "AQR Capital Management": "0001167557",
    "Renaissance Technologies": "0001037389",
}

# 모델 포트폴리오 (Core 30% + 5대 주도 섹터 Satellite 70%)
PORTFOLIO = [
    {"ticker": "QQQ",  "cusip": "46090E103", "name": "Invesco QQQ Trust",        "sector": "Core (지수 코어)",            "weight": 0.15, "eps_2027": "+14.5%"},
    {"ticker": "SPY",  "cusip": "78462F103", "name": "SPDR S&P 500 ETF",         "sector": "Core (지수 코어)",            "weight": 0.15, "eps_2027": "+11.2%"},
    {"ticker": "NVDA", "cusip": "67066G104", "name": "NVIDIA Corp",              "sector": "1. AI 반도체·광네트워킹",     "weight": 0.12, "eps_2027": "+28.4%"},
    {"ticker": "ANET", "cusip": "040413205", "name": "Arista Networks",          "sector": "1. AI 반도체·광네트워킹",     "weight": 0.08, "eps_2027": "+21.0%"},
    {"ticker": "LITE", "cusip": "55024U109", "name": "Lumentum Holdings",        "sector": "1. AI 반도체·광네트워킹",     "weight": 0.05, "eps_2027": "+42.0%"},
    {"ticker": "GEV",  "cusip": "36828A101", "name": "GE Vernova",               "sector": "2. 전력 인프라·방산",         "weight": 0.15, "eps_2027": "+34.5%"},
    {"ticker": "PLTR", "cusip": "69608A108", "name": "Palantir Technologies",    "sector": "3. AI 소프트웨어·사이버보안", "weight": 0.07, "eps_2027": "+31.0%"},
    {"ticker": "CRWD", "cusip": "22788C105", "name": "CrowdStrike Holdings",     "sector": "3. AI 소프트웨어·사이버보안", "weight": 0.05, "eps_2027": "+24.8%"},
    {"ticker": "LLY",  "cusip": "532457108", "name": "Eli Lilly and Co",         "sector": "4. 헬스케어·정밀진단",       "weight": 0.10, "eps_2027": "+25.2%"},
    {"ticker": "XOM",  "cusip": "30231G102", "name": "Exxon Mobil Corp",         "sector": "5. 에너지·인플레 헤지",       "weight": 0.08, "eps_2027": "+9.8%"},
]

# 5대 주도 업종별 5종목 (총 25종목) 관찰 유니버스
WATCHLIST_5_SECTORS = {
    "1. AI 네트워킹·광통신·스토리지": ["APH", "ANET", "LITE", "STX", "WDC"],
    "2. 산업재·방산·분산전원 인프라": ["GEV", "PWR", "RTX", "HWM", "BE"],
    "3. 사이버보안·AI 수익화 SW": ["CRWD", "PANW", "PLTR", "APP", "DDOG"],
    "4. 헬스케어·정밀진단·생명과학": ["LLY", "UNH", "TMO", "DHR", "NTRA"],
    "5. 전통 에너지·금광·방어 소비재": ["XOM", "COP", "AEM", "COST", "MNST"],
}
