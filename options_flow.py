# -*- coding: utf-8 -*-
"""
옵션 매매동향(Options Flow) 자동 수집 모듈
- Yahoo Finance 옵션 체인(근월물~차월물 4개 만기 합산) 기반:
  1) 콜/풋 거래량 비율 (Call/Put Volume Ratio)
  2) 콜/풋 미결제약정 비율 (Call/Put Open Interest Ratio)
  3) 이상 콜옵션 수급 스코어 (Unusual Call Activity: 거래량 > 미결제약정 종목 비중)
  4) 수급 시그널 판정 (강세 콜 집중 / 중립 / 풋 헤지 경계)
"""

import time
import pandas as pd
import yfinance as yf


def fetch_options_metrics(tickers: list[str]) -> pd.DataFrame:
    records = []
    for ticker in tickers:
        try:
            tk = yf.Ticker(ticker)
            expirations = tk.options[:4]  # 가까운 4개 만기일 옵션 체인 스캔
            if not expirations:
                raise ValueError("No options expirations")

            call_vol_total = 0.0
            put_vol_total = 0.0
            call_oi_total = 0.0
            put_oi_total = 0.0
            unusual_call_strikes = 0

            for exp in expirations:
                chain = tk.option_chain(exp)
                calls = chain.calls.fillna(0)
                puts = chain.puts.fillna(0)

                call_vol_total += float(calls["volume"].sum())
                put_vol_total += float(puts["volume"].sum())
                call_oi_total += float(calls["openInterest"].sum())
                put_oi_total += float(puts["openInterest"].sum())

                # 거래량이 기존 미결제약정(OI)을 초과하며 500계약 이상 터진 행사가 카운트
                unusual = calls[(calls["volume"] > calls["openInterest"]) & (calls["volume"] >= 500)]
                unusual_call_strikes += len(unusual)

            cp_vol_ratio = round(call_vol_total / max(put_vol_total, 1.0), 2)
            cp_oi_ratio = round(call_oi_total / max(put_oi_total, 1.0), 2)

            if cp_vol_ratio >= 1.5 and unusual_call_strikes >= 2:
                signal = "🔥 콜옵션 공격적 매수 유입"
            elif cp_vol_ratio >= 1.1:
                signal = "📈 콜 우위 (상방 베팅 지속)"
            elif cp_vol_ratio >= 0.8:
                signal = "⚖️ 중립 (양방향 균형)"
            else:
                signal = "🛡️ 풋 헤지 증가 (하방 경계)"

            records.append({
                "Ticker": ticker,
                "Call_Vol": int(call_vol_total),
                "Put_Vol": int(put_vol_total),
                "Call_Put_Vol_Ratio": cp_vol_ratio,
                "Call_Put_OI_Ratio": cp_oi_ratio,
                "Unusual_Call_Strikes": unusual_call_strikes,
                "Option_Signal": signal,
            })
        except Exception:
            records.append({
                "Ticker": ticker,
                "Call_Vol": 0,
                "Put_Vol": 0,
                "Call_Put_Vol_Ratio": 1.0,
                "Call_Put_OI_Ratio": 1.0,
                "Unusual_Call_Strikes": 0,
                "Option_Signal": "데이터 대기",
            })
        time.sleep(0.25)

    return pd.DataFrame(records)
