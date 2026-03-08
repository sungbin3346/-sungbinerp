# =============================================================================
# db.py — Supabase 연동 모듈
# =============================================================================
# Supabase PostgreSQL에 납품/결제 데이터를 영구 저장/조회합니다.
# 환경변수 SUPABASE_URL, SUPABASE_KEY 가 없으면 로컬 모드(세션만)로 동작합니다.
# =============================================================================

from __future__ import annotations
import os
import json
import pandas as pd
from datetime import datetime
from typing import Optional

# ── 환경변수 로드 (.env 파일 또는 Streamlit secrets) ─────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

try:
    import streamlit as st
    _SECRETS = st.secrets if hasattr(st, "secrets") else {}
except Exception:
    _SECRETS = {}

def _get_env(key: str) -> Optional[str]:
    """환경변수 → Streamlit secrets 순서로 값을 읽습니다."""
    val = os.environ.get(key)
    if not val and key in _SECRETS:
        val = _SECRETS[key]
    return val or None

SUPABASE_URL = _get_env("SUPABASE_URL")
SUPABASE_KEY = _get_env("SUPABASE_KEY")

# ── Supabase 클라이언트 초기화 ────────────────────────────────────────────────
_client = None

def get_client():
    """싱글톤 Supabase 클라이언트 반환. 환경변수 없으면 None."""
    global _client
    if _client is not None:
        return _client
    if not SUPABASE_URL or not SUPABASE_KEY:
        return None
    try:
        from supabase import create_client
        _client = create_client(SUPABASE_URL, SUPABASE_KEY)
        return _client
    except Exception:
        return None

def is_connected() -> bool:
    return get_client() is not None

# ── 테이블 이름 상수 ──────────────────────────────────────────────────────────
TABLE_SALES   = "erp_sales"
TABLE_PAYMENT = "erp_payment"

# =============================================================================
# Supabase SQL 스키마 (최초 1회 실행 필요 — README 참조)
# =============================================================================
SCHEMA_SQL = """
-- 납품 내역 테이블
CREATE TABLE IF NOT EXISTS erp_sales (
    id          BIGSERIAL PRIMARY KEY,
    납품일      DATE          NOT NULL,
    거래처      TEXT          NOT NULL DEFAULT '',
    교수        TEXT          NOT NULL DEFAULT '',
    연구원      TEXT          NOT NULL DEFAULT '',
    brand       TEXT          NOT NULL DEFAULT '',
    cat_no      TEXT          NOT NULL DEFAULT '',
    품명        TEXT          NOT NULL DEFAULT '',
    사이즈      TEXT          NOT NULL DEFAULT '',
    수량        INTEGER       NOT NULL DEFAULT 0,
    단가        BIGINT        NOT NULL DEFAULT 0,
    금액        BIGINT        NOT NULL DEFAULT 0,
    부가세포함  BIGINT        NOT NULL DEFAULT 0,
    월          TEXT          NOT NULL DEFAULT '',
    uploaded_at TIMESTAMPTZ   NOT NULL DEFAULT NOW()
);

-- 결제 내역 테이블
CREATE TABLE IF NOT EXISTS erp_payment (
    id          BIGSERIAL PRIMARY KEY,
    결제일      DATE          NOT NULL,
    교수        TEXT          NOT NULL DEFAULT '',
    연구원      TEXT          NOT NULL DEFAULT '',
    금액        BIGINT        NOT NULL DEFAULT 0,
    uploaded_at TIMESTAMPTZ   NOT NULL DEFAULT NOW()
);

-- 인덱스
CREATE INDEX IF NOT EXISTS idx_sales_거래처   ON erp_sales(거래처);
CREATE INDEX IF NOT EXISTS idx_sales_교수     ON erp_sales(교수);
CREATE INDEX IF NOT EXISTS idx_sales_월       ON erp_sales(월);
CREATE INDEX IF NOT EXISTS idx_payment_교수   ON erp_payment(교수);
"""

# =============================================================================
# 저장 함수
# =============================================================================

def _df_to_records(df: pd.DataFrame, table: str) -> list[dict]:
    """DataFrame → Supabase upsert용 dict 리스트 변환."""
    records = []
    for _, row in df.iterrows():
        if table == TABLE_SALES:
            records.append({
                "납품일":    row["납품일"].strftime("%Y-%m-%d") if pd.notna(row["납품일"]) else None,
                "거래처":   str(row.get("거래처", "")),
                "교수":     str(row.get("교수", "")),
                "연구원":   str(row.get("연구원", "")),
                "brand":    str(row.get("Brand", "")),
                "cat_no":   str(row.get("Cat.No", "")),
                "품명":     str(row.get("품명", "")),
                "사이즈":   str(row.get("사이즈", "")),
                "수량":     int(row.get("수량", 0)),
                "단가":     int(row.get("단가", 0)),
                "금액":     int(row.get("금액", 0)),
                "부가세포함": int(row.get("부가세포함", 0)),
                "월":       str(row.get("월", "")),
            })
        elif table == TABLE_PAYMENT:
            records.append({
                "결제일":   row["결제일"].strftime("%Y-%m-%d") if pd.notna(row["결제일"]) else None,
                "교수":     str(row.get("교수", "")),
                "연구원":   str(row.get("연구원", "")),
                "금액":     int(row.get("금액", 0)),
            })
    return records


def save_sales(df: pd.DataFrame) -> tuple[bool, str]:
    """납품 내역 DataFrame을 Supabase에 저장 (기존 데이터 전체 교체)."""
    client = get_client()
    if client is None:
        return False, "Supabase 미연결 — 세션에만 저장됩니다."
    try:
        # 기존 전체 삭제 후 재삽입 (upsert 대신 단순 replace)
        client.table(TABLE_SALES).delete().neq("id", 0).execute()
        records = _df_to_records(df, TABLE_SALES)
        # 500건씩 배치 삽입
        for i in range(0, len(records), 500):
            client.table(TABLE_SALES).insert(records[i:i+500]).execute()
        return True, f"납품 내역 {len(records):,}건 저장 완료"
    except Exception as e:
        return False, f"저장 실패: {e}"


def save_payment(df: pd.DataFrame) -> tuple[bool, str]:
    """결제 내역 DataFrame을 Supabase에 저장."""
    client = get_client()
    if client is None:
        return False, "Supabase 미연결 — 세션에만 저장됩니다."
    try:
        client.table(TABLE_PAYMENT).delete().neq("id", 0).execute()
        records = _df_to_records(df, TABLE_PAYMENT)
        for i in range(0, len(records), 500):
            client.table(TABLE_PAYMENT).insert(records[i:i+500]).execute()
        return True, f"결제 내역 {len(records):,}건 저장 완료"
    except Exception as e:
        return False, f"저장 실패: {e}"

# =============================================================================
# 조회 함수
# =============================================================================

def _records_to_sales_df(records: list[dict]) -> pd.DataFrame:
    """Supabase rows → app.py 호환 DataFrame 변환."""
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    # 컬럼명 원복
    rename_map = {"brand": "Brand", "cat_no": "Cat.No"}
    df = df.rename(columns=rename_map)
    # 날짜 변환
    df["납품일"] = pd.to_datetime(df["납품일"], errors="coerce")
    # 숫자 컬럼 int 변환
    for col in ["수량", "단가", "금액", "부가세포함"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)
    # 월 컬럼 보정
    if "월" not in df.columns or df["월"].eq("").all():
        df["월"] = df["납품일"].dt.to_period("M").astype(str)
    # 불필요 컬럼 제거
    for col in ["id", "uploaded_at"]:
        if col in df.columns:
            df = df.drop(columns=[col])
    return df


def _records_to_payment_df(records: list[dict]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df["결제일"] = pd.to_datetime(df["결제일"], errors="coerce")
    df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0).astype(int)
    for col in ["id", "uploaded_at"]:
        if col in df.columns:
            df = df.drop(columns=[col])
    return df


def load_sales() -> Optional[pd.DataFrame]:
    """Supabase에서 납품 내역 전체 조회. 없으면 None."""
    client = get_client()
    if client is None:
        return None
    try:
        # 1000건씩 페이지네이션
        all_records: list[dict] = []
        page = 0
        page_size = 1000
        while True:
            resp = (
                client.table(TABLE_SALES)
                .select("*")
                .range(page * page_size, (page + 1) * page_size - 1)
                .execute()
            )
            batch = resp.data or []
            all_records.extend(batch)
            if len(batch) < page_size:
                break
            page += 1
        if not all_records:
            return None
        return _records_to_sales_df(all_records)
    except Exception:
        return None


def load_payment() -> Optional[pd.DataFrame]:
    """Supabase에서 결제 내역 전체 조회. 없으면 None."""
    client = get_client()
    if client is None:
        return None
    try:
        all_records: list[dict] = []
        page = 0
        page_size = 1000
        while True:
            resp = (
                client.table(TABLE_PAYMENT)
                .select("*")
                .range(page * page_size, (page + 1) * page_size - 1)
                .execute()
            )
            batch = resp.data or []
            all_records.extend(batch)
            if len(batch) < page_size:
                break
            page += 1
        if not all_records:
            return None
        return _records_to_payment_df(all_records)
    except Exception:
        return None


def clear_all() -> tuple[bool, str]:
    """Supabase 전체 데이터 삭제."""
    client = get_client()
    if client is None:
        return False, "Supabase 미연결"
    try:
        client.table(TABLE_SALES).delete().neq("id", 0).execute()
        client.table(TABLE_PAYMENT).delete().neq("id", 0).execute()
        return True, "전체 데이터 삭제 완료"
    except Exception as e:
        return False, f"삭제 실패: {e}"
