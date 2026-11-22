# 미래에셋증권 일산WM - 헤지펀드(13F) · 옵션 수급 · 수익률 주간 자동화 앱

## 1. 핵심 기능
1. **수익률 자동 트래킹 (`update_weekly.py`)**: 직전 금요일 미국 마감 종가 기준 주간(1W) · 월간(1M) · 연초대비(YTD) · 2025년 이후 전체 누적 수익률 산출
2. **옵션 매매동향 자동 스캔 (`options_flow.py`)**: 근월물~차월물 4개 만기 옵션 체인을 스캔해 **콜/풋 거래량 비율(Call/Put Volume Ratio)**, **콜/풋 미결제약정 비율**, **이상 콜옵션 급증 행사가 수(Unusual Call Activity)** 및 수급 시그널 자동 판정
3. **헤지펀드 13F 공시 자동 추적 (`sec_13f.py`)**: 미국 SEC EDGAR 공식 API에서 대형 8대 헤지펀드/마켓메이커(Citadel, Millennium, Point72, Jane Street, SIG, Bridgewater, AQR, Renaissance)의 최신 13F 공시를 파싱하여 **종목별 보유 펀드 수 및 13F 콜옵션 보유 현황** 자동 집계
4. **GitHub Actions 매주 일요일 무인 자동화 (`.github/workflows/weekly_sunday_update.yml`)**: 매주 일요일 오전 9시(KST) 자동 실행 후 Excel 누적 로그와 PPTX 리포트를 저장소에 자동 커밋

## 2. GitHub 연동 및 매주 일요일 자동 실행 세팅 (5분 완성)
1. GitHub에서 새 저장소(Repository)를 생성합니다 (예: `ilsan-wm-us-flow-tracker`).
2. 압축을 푼 폴더 내 모든 파일(숨김 폴더 `.github` 포함)을 해당 GitHub 저장소에 업로드(Push)합니다.
3. 저장소 상단 **[Settings] → [Actions] → [General] → Workflow permissions**에서 **`Read and write permissions`**를 선택하고 Save를 누릅니다.
4. 상단 **[Actions]** 탭에서 **`Weekly Sunday HedgeFund & Options Flow Update`** 워크플로우를 선택하고 **[Run workflow]** 버튼을 한 번 눌러 첫 실행을 테스트합니다.
5. 이후 매주 일요일 오전 9시(KST)마다 자동으로 실행되며, Streamlit Cloud(`share.streamlit.io`)에 이 GitHub 저장소를 연결하면 스마트폰·PC 어디서든 웹 대시보드로 접속할 수 있습니다.
