# =============================================================================
# sheets_db.py — Google Sheets 영구 저장 모듈
# =============================================================================
# Streamlit Cloud Secrets 에 [gcp_service_account] 및 SPREADSHEET_ID 를
# 설정하면 활성화됩니다. 없으면 세션 전용 모드로 동작합니다.
#
# Secrets 예시 (.streamlit/secrets.toml):
#   [gcp_service_account]
#   type = "service_account"
#   project_id = "..."
#   private_key_id = "..."
#   private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n"
#   client_email = "xxx@project.iam.gserviceaccount.com"
#   ...
#
#   SPREADSHEET_ID = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgVE2upms"
# =============================================================================

from __future__ import annotations
import pandas as pd
import streamlit as st
from typing import Optional

# ── 시트 이름 상수 ─────────────────────────────────────────────────────────────
SHEET_SALES   = "납품내역"
SHEET_PAYMENT = "결제내역"

# ── 날짜 포맷 ──────────────────────────────────────────────────────────────────
DATE_FMT = "%Y-%m-%d"

# ── Streamlit Cloud 영구 저장 연결 여부 확인 ──────────────────────────────────
def _get_client():
    """
    gspread 클라이언트 반환.
    Streamlit Secrets 에 [gcp_service_account] 없으면 None 반환.
    """
    try:
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        return gspread.authorize(creds)
    except Exception:
        return None


def _get_spreadsheet():
    """스프레드시트 객체 반환. 실패 시 None."""
    try:
        gc = _get_client()
        if gc is None:
            return None
        spreadsheet_id = st.secrets.get("SPREADSHEET_ID", "")
        if not spreadsheet_id:
            return None
        return gc.open_by_key(spreadsheet_id)
    except Exception:
        return None


def is_connected() -> bool:
    """Google Sheets 연결 가능 여부"""
    try:
        gc = _get_client()
        sid = st.secrets.get("SPREADSHEET_ID", "")
        return gc is not None and bool(sid)
    except Exception:
        return False


# =============================================================================
# 워크시트 초기화 (없으면 생성)
# =============================================================================

def _get_or_create_sheet(ss, sheet_name: str, headers: list[str]):
    """
    시트가 없으면 생성하고 헤더를 씁니다.
    있으면 그대로 반환합니다.
    """
    try:
        ws = ss.worksheet(sheet_name)
        # 헤더 없으면 추가
        existing = ws.row_values(1)
        if not existing:
            ws.insert_row(headers, index=1)
        return ws
    except Exception:
        ws = ss.add_worksheet(title=sheet_name, rows=5000, cols=len(headers))
        ws.insert_row(headers, index=1)
        return ws


# =============================================================================
# 저장 함수
# =============================================================================

def save_sales(df: pd.DataFrame) -> tuple[bool, str]:
    """
    납품 내역 DataFrame 을 Google Sheets 납품내역 시트에 저장합니다.
    기존 데이터 전체를 삭제하고 재삽입합니다.
    """
    ss = _get_spreadsheet()
    if ss is None:
        return False, "Google Sheets 미연결 — 세션에만 저장됩니다."
    try:
        headers = [
            "납품일", "거래처", "교수", "연구원", "Brand", "Cat.No",
            "품명", "사이즈", "수량", "단가", "금액", "부가세포함", "월"
        ]
        ws = _get_or_create_sheet(ss, SHEET_SALES, headers)

        # 헤더 행(1행) 이후 전체 삭제
        ws.resize(rows=1)

        # 데이터 배열 준비
        rows = []
        for _, r in df.iterrows():
            date_val = (
                r["납품일"].strftime(DATE_FMT)
                if pd.notna(r.get("납품일")) else ""
            )
            rows.append([
                date_val,
                str(r.get("거래처", "")),
                str(r.get("교수", "")),
                str(r.get("연구원", "")),
                str(r.get("Brand", "")),
                str(r.get("Cat.No", "")),
                str(r.get("품명", "")),
                str(r.get("사이즈", "")),
                int(r.get("수량", 0)),
                int(r.get("단가", 0)),
                int(r.get("금액", 0)),
                int(r.get("부가세포함", 0)),
                str(r.get("월", "")),
            ])

        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")

        return True, f"납품 내역 {len(rows):,}건 Google Sheets 저장 완료"
    except Exception as e:
        return False, f"저장 실패: {e}"


def save_payment(df: pd.DataFrame) -> tuple[bool, str]:
    """결제 내역 DataFrame 을 Google Sheets 결제내역 시트에 저장합니다."""
    ss = _get_spreadsheet()
    if ss is None:
        return False, "Google Sheets 미연결 — 세션에만 저장됩니다."
    try:
        headers = ["결제일", "교수", "연구원", "금액"]
        ws = _get_or_create_sheet(ss, SHEET_PAYMENT, headers)

        ws.resize(rows=1)

        rows = []
        for _, r in df.iterrows():
            date_val = (
                r["결제일"].strftime(DATE_FMT)
                if pd.notna(r.get("결제일")) else ""
            )
            rows.append([
                date_val,
                str(r.get("교수", "")),
                str(r.get("연구원", "")),
                int(r.get("금액", 0)),
            ])

        if rows:
            ws.append_rows(rows, value_input_option="USER_ENTERED")

        return True, f"결제 내역 {len(rows):,}건 Google Sheets 저장 완료"
    except Exception as e:
        return False, f"저장 실패: {e}"


# =============================================================================
# 조회 함수
# =============================================================================

def load_sales() -> Optional[pd.DataFrame]:
    """Google Sheets 납품내역 시트에서 전체 데이터 로드."""
    ss = _get_spreadsheet()
    if ss is None:
        return None
    try:
        ws = ss.worksheet(SHEET_SALES)
        all_values = ws.get_all_records(
            expected_headers=[
                "납품일", "거래처", "교수", "연구원", "Brand", "Cat.No",
                "품명", "사이즈", "수량", "단가", "금액", "부가세포함", "월"
            ]
        )
        if not all_values:
            return None

        df = pd.DataFrame(all_values)

        # 날짜 변환
        df["납품일"] = pd.to_datetime(df["납품일"], errors="coerce")
        df = df.dropna(subset=["납품일"])

        # 숫자 컬럼 변환
        for col in ["수량", "단가", "금액", "부가세포함"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

        # 문자열 정리
        for col in ["거래처", "교수", "연구원", "Brand", "Cat.No", "품명", "사이즈"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        # 월 컬럼 보정
        if "월" not in df.columns or df["월"].eq("").all():
            df["월"] = df["납품일"].dt.to_period("M").astype(str)

        return df
    except Exception:
        return None


def load_payment() -> Optional[pd.DataFrame]:
    """Google Sheets 결제내역 시트에서 전체 데이터 로드."""
    ss = _get_spreadsheet()
    if ss is None:
        return None
    try:
        ws = ss.worksheet(SHEET_PAYMENT)
        all_values = ws.get_all_records(
            expected_headers=["결제일", "교수", "연구원", "금액"]
        )
        if not all_values:
            return None

        df = pd.DataFrame(all_values)
        df["결제일"] = pd.to_datetime(df["결제일"], errors="coerce")
        df = df.dropna(subset=["결제일"])
        df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0).astype(int)

        for col in ["교수", "연구원"]:
            if col in df.columns:
                df[col] = df[col].fillna("").astype(str).str.strip()

        return df
    except Exception:
        return None


def clear_all() -> tuple[bool, str]:
    """Google Sheets 납품내역 / 결제내역 전체 삭제."""
    ss = _get_spreadsheet()
    if ss is None:
        return False, "Google Sheets 미연결"
    try:
        msgs = []
        for sheet_name in [SHEET_SALES, SHEET_PAYMENT]:
            try:
                ws = ss.worksheet(sheet_name)
                ws.resize(rows=1)   # 헤더(1행)만 남기고 삭제
                msgs.append(f"{sheet_name} 초기화")
            except Exception:
                pass
        return True, " / ".join(msgs) + " 완료"
    except Exception as e:
        return False, f"삭제 실패: {e}"
