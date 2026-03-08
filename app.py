# =============================================================================
# 생명공학 시약 기술영업 개인용 ERP - app.py
# =============================================================================
# 작성자: AI Assistant
# 설명: 납품 내역 & 결제 내역 엑셀 파일을 업로드하여
#       매출 현황을 시각화하고 교수(고객)별 미수금을 관리하는 Streamlit 대시보드
# 실행: streamlit run app.py
# =============================================================================

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import warnings

# Google Sheets 영구 저장 모듈 (Secrets 없으면 세션 전용 모드)
try:
    import sheets_db as erp_db
    _DB_MODULE_OK = True
except Exception:
    _DB_MODULE_OK = False

warnings.filterwarnings("ignore")

# =============================================================================
# 1. 페이지 전역 설정
# =============================================================================
st.set_page_config(
    page_title="영업 관리 ERP",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# 2. 전역 CSS 스타일 정의
# =============================================================================
st.markdown(
    """
    <style>
    /* ====================================================
       전체 레이아웃
    ==================================================== */
    .main {
        background-color: #f0f2f6;
    }
    /* 메인 콘텐츠 영역 기본 텍스트 */
    .main p, .main span, .main div, .main label,
    .block-container p, .block-container span,
    .block-container div, .block-container label {
        color: #1a1a2e;
    }

    /* ====================================================
       사이드바 — 다크 네이비 배경, 모든 텍스트 강제 밝게
    ==================================================== */
    section[data-testid="stSidebar"] {
        background-color: #1e2235 !important;
    }
    /* 사이드바 내 모든 일반 텍스트 */
    section[data-testid="stSidebar"] p,
    section[data-testid="stSidebar"] span,
    section[data-testid="stSidebar"] div,
    section[data-testid="stSidebar"] label,
    section[data-testid="stSidebar"] h1,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3,
    section[data-testid="stSidebar"] small,
    section[data-testid="stSidebar"] li {
        color: #e8eaf0 !important;
    }
    /* 사이드바 파일 업로더 — 내부 배경·텍스트 재정의 */
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] {
        background-color: #2a3050 !important;
        border: 1.5px dashed #5b6fa6 !important;
        border-radius: 10px !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] * {
        color: #c8d0e7 !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button {
        background-color: #3d4f88 !important;
        color: #ffffff !important;
        border: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stFileUploader"] button:hover {
        background-color: #4f65b0 !important;
    }
    /* 사이드바 버튼 */
    section[data-testid="stSidebar"] button {
        background-color: #3a4268 !important;
        color: #ffffff !important;
        border: 1px solid #5b6fa6 !important;
    }
    section[data-testid="stSidebar"] button:hover {
        background-color: #4f5c8a !important;
    }
    /* 사이드바 구분선 */
    section[data-testid="stSidebar"] hr {
        border-color: #3a4268 !important;
    }
    /* 사이드바 다운로드 버튼 */
    section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button {
        background-color: #2e7d52 !important;
        color: #ffffff !important;
        border: none !important;
    }
    section[data-testid="stSidebar"] [data-testid="stDownloadButton"] button:hover {
        background-color: #3a9e68 !important;
    }

    /* ====================================================
       KPI 카드
    ==================================================== */
    .kpi-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 18px 22px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.09);
        border-left: 5px solid #4F8BF9;
        margin-bottom: 8px;
    }
    .kpi-label {
        font-size: 12px;
        color: #5a6478 !important;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 6px;
    }
    .kpi-value {
        font-size: 26px;
        font-weight: 800;
        color: #1a1a2e !important;
        line-height: 1.2;
    }
    .kpi-sub {
        font-size: 11px;
        color: #8892a4 !important;
        margin-top: 5px;
    }

    /* ====================================================
       미수금 알림 배너 카드
    ==================================================== */
    .alert-danger {
        background: linear-gradient(135deg, #ff6b6b, #c0392b);
        color: #ffffff !important;
        border-radius: 12px;
        padding: 22px 28px;
        box-shadow: 0 4px 18px rgba(192,57,43,0.35);
        text-align: center;
    }
    .alert-success {
        background: linear-gradient(135deg, #00b894, #00796b);
        color: #ffffff !important;
        border-radius: 12px;
        padding: 22px 28px;
        box-shadow: 0 4px 18px rgba(0,121,107,0.3);
        text-align: center;
    }
    .alert-value {
        font-size: 30px;
        font-weight: 900;
        letter-spacing: -1px;
        color: #ffffff !important;
    }
    .alert-label {
        font-size: 13px;
        color: rgba(255,255,255,0.92) !important;
        margin-top: 5px;
    }

    /* ====================================================
       섹션 헤더
    ==================================================== */
    .section-header {
        font-size: 17px;
        font-weight: 700;
        color: #1a1a2e !important;
        border-bottom: 2.5px solid #4F8BF9;
        padding-bottom: 6px;
        margin: 22px 0 14px 0;
    }

    /* ====================================================
       업로드 상태 배지
    ==================================================== */
    .badge-success {
        display: inline-block;
        background: #00b894;
        color: #ffffff !important;
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 12px;
        font-weight: 700;
        margin-top: 4px;
    }
    .badge-warning {
        display: inline-block;
        background: #f39c12;
        color: #ffffff !important;
        border-radius: 20px;
        padding: 4px 14px;
        font-size: 12px;
        font-weight: 700;
        margin-top: 4px;
    }

    /* ====================================================
       메인 탭 스타일
    ==================================================== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
        background-color: #e8ecf4;
        padding: 6px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        font-weight: 600;
        font-size: 14px;
        color: #4a5568 !important;
        padding: 8px 20px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #ffffff !important;
        color: #1a1a2e !important;
        box-shadow: 0 2px 6px rgba(0,0,0,0.1);
    }

    /* ====================================================
       DataFrame / 테이블
    ==================================================== */
    .stDataFrame {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #dde3ee;
    }

    /* ====================================================
       메인 타이틀 영역
    ==================================================== */
    .main-title {
        font-size: 30px;
        font-weight: 900;
        color: #1a1a2e !important;
        letter-spacing: -0.5px;
    }
    .main-subtitle {
        font-size: 14px;
        color: #5a6478 !important;
        margin-top: 4px;
    }

    /* ====================================================
       빈 상태 안내 카드
    ==================================================== */
    .empty-state {
        text-align: center;
        padding: 80px 20px;
        background: #ffffff;
        border-radius: 16px;
        border: 2px dashed #c8d0e0;
        margin: 20px 0;
    }
    .empty-state .icon { font-size: 56px; margin-bottom: 12px; }
    .empty-state h3 { color: #5a6478 !important; font-size: 20px; font-weight: 700; }
    .empty-state p { color: #8892a4 !important; font-size: 14px; line-height: 1.6; }

    /* ====================================================
       Streamlit 기본 컴포넌트 재정의
    ==================================================== */
    /* metric 카드 */
    [data-testid="metric-container"] {
        background: #ffffff;
        border-radius: 10px;
        padding: 14px 18px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
    }
    [data-testid="metric-container"] label,
    [data-testid="stMetricLabel"] > div {
        color: #5a6478 !important;
        font-weight: 600 !important;
    }
    [data-testid="stMetricValue"] > div {
        color: #1a1a2e !important;
        font-weight: 800 !important;
    }
    [data-testid="stMetricDelta"] > div { font-weight: 600 !important; }

    /* expander 헤더 */
    .streamlit-expanderHeader {
        background: #ffffff !important;
        color: #1a1a2e !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }
    .streamlit-expanderHeader:hover {
        background: #f0f2f6 !important;
    }

    /* selectbox, dropdown 텍스트 */
    [data-testid="stSelectbox"] label { color: #1a1a2e !important; font-weight: 600; }
    [data-baseweb="select"] div { color: #1a1a2e !important; }

    /* 캡션 텍스트 */
    .stCaption, caption { color: #5a6478 !important; }

    /* info / warning / error 박스 */
    .stAlert p { color: #1a1a2e !important; }

    /* 메인 구분선 */
    hr { border-color: #dde3ee !important; }

    /* ====================================================
       모바일 / iPhone 반응형 디자인
    ==================================================== */
    /* 수평 스크롤 방지 */
    html, body { overflow-x: hidden !important; }

    /* iOS safe-area 인셋 지원 */
    .main .block-container {
        padding-left: max(1rem, env(safe-area-inset-left)) !important;
        padding-right: max(1rem, env(safe-area-inset-right)) !important;
        padding-bottom: max(1rem, env(safe-area-inset-bottom)) !important;
    }

    /* 이미지/테이블 최대 너비 */
    img, table { max-width: 100% !important; }

    /* 최소 폰트 크기 16px (iOS 자동 확대 방지) */
    input, select, textarea {
        font-size: 16px !important;
        -webkit-text-size-adjust: 100%;
    }

    /* 터치 스크롤 부드럽게 */
    * { -webkit-overflow-scrolling: touch; }

    /* 300ms 클릭 딜레이 제거 */
    a, button, [role="button"] { touch-action: manipulation; }

    /* 태블릿 (768px 이하) */
    @media (max-width: 768px) {
        .main-title { font-size: 20px !important; }
        .main-subtitle { font-size: 12px !important; }
        .kpi-value { font-size: 18px !important; }
        .kpi-label { font-size: 11px !important; }
        .kpi-card { padding: 12px 14px !important; }
        .section-header { font-size: 14px !important; }
        .alert-value { font-size: 22px !important; }

        /* 탭 폰트 축소 */
        .stTabs [data-baseweb="tab"] {
            font-size: 12px !important;
            padding: 6px 10px !important;
        }
    }

    /* 모바일 (480px 이하) — 편집 버튼 숨김 (읽기 전용) */
    @media (max-width: 480px) {
        .main-title { font-size: 17px !important; }
        .kpi-value { font-size: 16px !important; }

        /* 데이터프레임 가로 스크롤 허용 */
        .stDataFrame { overflow-x: auto !important; }

        /* 사이드바 숨겨진 상태에서 메인 영역 패딩 최소화 */
        .block-container { padding: 8px !important; }

        /* ── 모바일 읽기 전용: 업로드·초기화·다운로드 영역 숨김 ── */
        /* 파일 업로드 위젯 */
        section[data-testid="stSidebar"] [data-testid="stFileUploader"] { display: none !important; }
        /* 데이터 초기화 버튼 */
        section[data-testid="stSidebar"] button { display: none !important; }
        /* 사이드바 다운로드 버튼 */
        section[data-testid="stSidebar"] [data-testid="stDownloadButton"] { display: none !important; }
        /* 조회 탭의 필터 초기화 버튼 */
        [data-testid="stMainBlockContainer"] button[kind="secondary"] { display: none !important; }
        /* 엑셀 다운로드 버튼 */
        [data-testid="stDownloadButton"] { display: none !important; }
        /* 사이드바 섹션 헤더 (업로드 섹션) */
        section[data-testid="stSidebar"] .stMarkdown h3:nth-of-type(-n+3) { display: none !important; }
    }

    /* Supabase 연결 상태 배지 */
    .db-connected {
        display: inline-flex; align-items: center; gap: 6px;
        background: #e8f8f0; color: #1e8449;
        border: 1px solid #27ae60; border-radius: 20px;
        padding: 4px 12px; font-size: 12px; font-weight: 700;
    }
    .db-disconnected {
        display: inline-flex; align-items: center; gap: 6px;
        background: #fef9e7; color: #b7950b;
        border: 1px solid #f39c12; border-radius: 20px;
        padding: 4px 12px; font-size: 12px; font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# PWA 메타태그 + Viewport 주입 (iPhone 홈 화면 추가 지원)
st.markdown(
    """
    <meta name="viewport" content="width=device-width, initial-scale=1.0,
          maximum-scale=1.0, user-scalable=no, viewport-fit=cover">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="영업ERP">
    <meta name="mobile-web-app-capable" content="yes">
    <meta name="theme-color" content="#1e2235">
    <link rel="manifest" href="/app/static/manifest.json">
    <link rel="apple-touch-icon" href="/app/static/apple-touch-icon.png">
    <style>
    /* iOS 상태바 영역 보호 */
    @supports (padding-top: env(safe-area-inset-top)) {
        .stApp {
            padding-top: env(safe-area-inset-top) !important;
        }
    }
    </style>
    <script>
    // Service Worker 등록 (PWA 오프라인 지원)
    if ('serviceWorker' in navigator) {
        window.addEventListener('load', function() {
            navigator.serviceWorker.register('/app/static/sw.js', {scope: '/'})
                .then(function(reg) { console.log('SW 등록 성공:', reg.scope); })
                .catch(function(err) { console.log('SW 등록 실패:', err); });
        });
    }
    </script>
    """,
    unsafe_allow_html=True,
)

# =============================================================================
# 3. 유틸리티 함수 정의
# =============================================================================

def format_currency(value: int | float) -> str:
    """금액을 ₩1,234,567 형식으로 포맷팅"""
    try:
        return f"₩{int(value):,}"
    except (ValueError, TypeError):
        return "₩0"


def format_currency_short(value: int | float) -> str:
    """금액을 억/만 단위 축약 형식으로 포맷팅 (KPI용)"""
    try:
        v = int(value)
        if abs(v) >= 100_000_000:
            return f"₩{v / 100_000_000:.1f}억"
        elif abs(v) >= 10_000:
            return f"₩{v / 10_000:.0f}만"
        else:
            return f"₩{v:,}"
    except (ValueError, TypeError):
        return "₩0"


def load_sales_data(file) -> pd.DataFrame:
    """
    납품 내역 엑셀 파일을 로드하고 전처리합니다.
    - 날짜 컬럼 datetime 변환
    - NaN → 0 처리
    - 금액 컬럼 int 변환
    """
    required_cols = ["납품일", "거래처", "교수", "연구원", "Brand", "Cat.No",
                     "품명", "사이즈", "수량", "단가", "금액", "부가세포함"]

    df = pd.read_excel(file, engine="openpyxl")

    # 컬럼명 공백 제거
    df.columns = df.columns.str.strip()

    # 필수 컬럼 존재 여부 검증
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"납품 내역에 필수 컬럼이 없습니다: {missing}")

    # 날짜 변환
    df["납품일"] = pd.to_datetime(df["납품일"], errors="coerce")

    # 문자열 컬럼 NaN → 빈 문자열
    str_cols = ["거래처", "교수", "연구원", "Brand", "Cat.No", "품명", "사이즈"]
    for col in str_cols:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # 숫자 컬럼 NaN → 0, int 변환 (부동소수점 오차 제거)
    num_cols = ["수량", "단가", "금액", "부가세포함"]
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    # 유효하지 않은 날짜 행 제거
    df = df.dropna(subset=["납품일"])

    # 월별 컬럼 추가 (차트용)
    df["월"] = df["납품일"].dt.to_period("M").astype(str)

    return df


def load_payment_data(file) -> pd.DataFrame:
    """
    결제 내역 엑셀 파일을 로드하고 전처리합니다.
    - 날짜 컬럼 datetime 변환
    - NaN → 0 처리
    - 금액 컬럼 int 변환
    """
    required_cols = ["결제일", "교수", "연구원", "금액"]

    df = pd.read_excel(file, engine="openpyxl")

    # 컬럼명 공백 제거
    df.columns = df.columns.str.strip()

    # 필수 컬럼 존재 여부 검증
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"결제 내역에 필수 컬럼이 없습니다: {missing}")

    # 날짜 변환
    df["결제일"] = pd.to_datetime(df["결제일"], errors="coerce")

    # 문자열 컬럼
    for col in ["교수", "연구원"]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    # 숫자 컬럼 NaN → 0, int 변환
    df["금액"] = pd.to_numeric(df["금액"], errors="coerce").fillna(0).astype(int)

    # 유효하지 않은 날짜 행 제거
    df = df.dropna(subset=["결제일"])

    return df


def render_kpi_card(label: str, value: str, sub: str = "", color: str = "#4F8BF9"):
    """KPI 카드 HTML 렌더링"""
    st.markdown(
        f"""
        <div class="kpi-card" style="border-left-color: {color};">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =============================================================================
# 4. 사이드바: 파일 업로드 & 세션 관리
# =============================================================================

with st.sidebar:
    st.markdown("## 🧬 영업 관리 ERP")
    st.markdown("---")

    # ── 납품 내역 업로드 ──────────────────────────────────────────────────
    st.markdown("### 📦 납품 내역 업로드")
    sales_file = st.file_uploader(
        "납품 내역 엑셀 파일",
        type=["xlsx", "xls"],
        key="sales_uploader",
        help="필수 컬럼: 납품일, 거래처, 교수, 연구원, Brand, Cat.No, 품명, 사이즈, 수량, 단가, 금액, 부가세포함",
    )

    # ── 납품 파일이 업로드되면 세션에 로드
    if sales_file is not None:
        try:
            df_sales = load_sales_data(sales_file)
            st.session_state["df_sales"] = df_sales
            st.markdown(
                f'<span class="badge-success">✅ {len(df_sales):,}건 로드 완료</span>',
                unsafe_allow_html=True,
            )
            # Supabase 자동 저장
            if _DB_MODULE_OK and erp_db.is_connected():
                ok, msg = erp_db.save_sales(df_sales)
                if ok:
                    st.success(f"📊 {msg}")
                else:
                    st.warning(f"📊 {msg}")
        except Exception as e:
            st.error(f"납품 내역 오류: {e}")
            st.session_state.pop("df_sales", None)
    elif "df_sales" in st.session_state:
        st.markdown(
            f'<span class="badge-success">✅ {len(st.session_state["df_sales"]):,}건 유지 중</span>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<span class="badge-warning">⚠️ 파일 미업로드</span>', unsafe_allow_html=True)

    st.markdown("---")

    # ── 결제 내역 업로드 ──────────────────────────────────────────────────
    st.markdown("### 💳 결제 내역 업로드")
    payment_file = st.file_uploader(
        "결제 내역 엑셀 파일",
        type=["xlsx", "xls"],
        key="payment_uploader",
        help="필수 컬럼: 결제일, 교수, 연구원, 금액",
    )

    if payment_file is not None:
        try:
            df_payment = load_payment_data(payment_file)
            st.session_state["df_payment"] = df_payment
            st.markdown(
                f'<span class="badge-success">✅ {len(df_payment):,}건 로드 완료</span>',
                unsafe_allow_html=True,
            )
            # Supabase 자동 저장
            if _DB_MODULE_OK and erp_db.is_connected():
                ok, msg = erp_db.save_payment(df_payment)
                if ok:
                    st.success(f"📊 {msg}")
                else:
                    st.warning(f"📊 {msg}")
        except Exception as e:
            st.error(f"결제 내역 오류: {e}")
            st.session_state.pop("df_payment", None)
    elif "df_payment" in st.session_state:
        st.markdown(
            f'<span class="badge-success">✅ {len(st.session_state["df_payment"]):,}건 유지 중</span>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<span class="badge-warning">⚠️ 파일 미업로드</span>', unsafe_allow_html=True)

    st.markdown("---")

    # ── 데이터 초기화 버튼 ─────────────────────────────────────────────────
    st.markdown("### ⚙️ 설정")
    if st.button("🗑️ 데이터 초기화", width="stretch", type="secondary"):
        # 세션 스테이트에서 데이터 삭제
        for key in ["df_sales", "df_payment"]:
            st.session_state.pop(key, None)
        # Supabase 데이터도 삭제
        if _DB_MODULE_OK and erp_db.is_connected():
            erp_db.clear_all()
        st.success("데이터가 초기화되었습니다.")
        st.rerun()

    # ── Supabase 연결 상태 표시 ───────────────────────────────────────────
    st.markdown("---")
    if _DB_MODULE_OK and erp_db.is_connected():
        st.markdown(
            '<span class="db-connected">📊 Google Sheets 연결됨 — 영구 저장 활성</span>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div style="font-size:11px; color:#b0bcd4; margin-top:6px;">'
            '데이터가 Google Sheets에 자동 저장됩니다.<br>앱 재시작 후에도 데이터 유지.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<span class="db-disconnected">💾 로컬 세션 모드</span>',
            unsafe_allow_html=True,
        )
        st.markdown(
            '<div style="font-size:11px; color:#b0bcd4; margin-top:6px;">'
            'Streamlit Secrets에 Google Sheets<br>'
            '서비스 계정을 설정하면 영구 저장됩니다.</div>',
            unsafe_allow_html=True,
        )

    # ── 사이드바 하단 정보 ─────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        """
        <div style="font-size:12px; color:#b0bcd4; line-height:1.8; background:#252b45;
                    border-radius:10px; padding:14px 16px; margin-top:4px;">
        📌 <b style="color:#e0e6f5;">사용 방법</b><br>
        1. 납품 내역 엑셀 업로드<br>
        2. 결제 내역 엑셀 업로드<br>
        3. 대시보드/장부 탭 확인<br><br>
        📊 <b style="color:#e0e6f5;">영구 저장 (Google Sheets)</b><br>
        Streamlit Cloud Secrets에<br>
        서비스 계정을 설정하면<br>
        앱 재시작 후에도 데이터 유지.<br><br>
        📱 <b style="color:#e0e6f5;">모바일</b><br>
        아이폰에서는 읽기 전용 모드.<br>
        PC에서만 업로드/편집 가능.
        </div>
        """,
        unsafe_allow_html=True,
    )

# =============================================================================
# 5. 메인 화면: 타이틀
# =============================================================================

st.markdown(
    """
    <div style="padding: 8px 0 4px 0;">
        <span class="main-title">🧬 생명공학 시약 영업 관리 ERP</span><br>
        <span class="main-subtitle">납품 내역과 결제 내역을 업로드하면 매출 현황과 미수금을 자동으로 분석합니다.</span>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("---")

# =============================================================================
# 6. 세션 상태에서 데이터 불러오기
# =============================================================================

# ── Google Sheets 자동 로드 (세션에 데이터 없고 Sheets 연결된 경우) ─────────
if _DB_MODULE_OK and erp_db.is_connected():
    if "df_sales" not in st.session_state:
        with st.spinner("📊 Google Sheets에서 납품 내역 불러오는 중..."):
            _loaded_sales = erp_db.load_sales()
            if _loaded_sales is not None and not _loaded_sales.empty:
                st.session_state["df_sales"] = _loaded_sales
    if "df_payment" not in st.session_state:
        with st.spinner("📊 Google Sheets에서 결제 내역 불러오는 중..."):
            _loaded_payment = erp_db.load_payment()
            if _loaded_payment is not None and not _loaded_payment.empty:
                st.session_state["df_payment"] = _loaded_payment

has_sales = "df_sales" in st.session_state
has_payment = "df_payment" in st.session_state

if has_sales:
    df_sales = st.session_state["df_sales"]
if has_payment:
    df_payment = st.session_state["df_payment"]

# =============================================================================
# 7. 탭 레이아웃 구성
# =============================================================================
tab1, tab2, tab3 = st.tabs([
    "📊 대시보드 (Dashboard)",
    "🔍 납품 조회 (Inquiry)",
    "📒 장부 관리 (Ledger)",
])


# =============================================================================
# Tab 1: 대시보드
# =============================================================================
with tab1:
    if not has_sales:
        # 업로드 안내 화면
        st.markdown(
            """
            <div class="empty-state">
                <div class="icon">📦</div>
                <h3>납품 내역을 업로드해주세요</h3>
                <p>왼쪽 사이드바에서 납품 내역 엑셀 파일을 업로드하면<br>
                대시보드가 자동으로 생성됩니다.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        # ── KPI 지표 ───────────────────────────────────────────────────────
        st.markdown('<div class="section-header">📈 핵심 성과 지표 (KPI)</div>', unsafe_allow_html=True)

        total_revenue = int(df_sales["부가세포함"].sum())
        total_transactions = len(df_sales)
        total_clients = df_sales["거래처"].nunique()
        total_professors = df_sales["교수"].replace("", pd.NA).dropna().nunique()
        avg_per_transaction = total_revenue // total_transactions if total_transactions > 0 else 0

        col_k1, col_k2, col_k3, col_k4 = st.columns(4)
        with col_k1:
            render_kpi_card(
                "총 매출액 (부가세 포함)",
                format_currency_short(total_revenue),
                f"정확한 금액: {format_currency(total_revenue)}",
                "#4F8BF9",
            )
        with col_k2:
            render_kpi_card(
                "총 거래 건수",
                f"{total_transactions:,}건",
                f"건당 평균: {format_currency(avg_per_transaction)}",
                "#00b894",
            )
        with col_k3:
            render_kpi_card(
                "거래처 수",
                f"{total_clients:,}곳",
                "등록된 거래처 (기관) 수",
                "#fd79a8",
            )
        with col_k4:
            render_kpi_card(
                "담당 교수 수",
                f"{total_professors:,}명",
                "납품 내역 기준 교수 수",
                "#e17055",
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # ── 차트 1: 월별 매출 추이 ─────────────────────────────────────────
        st.markdown('<div class="section-header">📅 월별 매출 추이</div>', unsafe_allow_html=True)

        monthly_sales = (
            df_sales.groupby("월")["부가세포함"]
            .sum()
            .reset_index()
            .rename(columns={"부가세포함": "매출액"})
        )
        # 월 컬럼을 반드시 문자열로 변환 후 정렬 (datetime 오인 방지)
        monthly_sales["월"] = monthly_sales["월"].astype(str)
        monthly_sales = monthly_sales.sort_values("월")
        months_sorted = monthly_sales["월"].tolist()  # 정렬된 월 목록 (category_orders용)

        # 포맷팅된 레이블 추가 (호버용)
        monthly_sales["매출액_표시"] = monthly_sales["매출액"].apply(format_currency)

        fig_monthly = px.bar(
            monthly_sales,
            x="월",
            y="매출액",
            text="매출액_표시",
            color="매출액",
            color_continuous_scale=["#a8d8ea", "#4F8BF9", "#1a1a9e"],
            labels={"매출액": "매출액 (원)", "월": ""},
            title="",
            category_orders={"월": months_sorted},  # X축 순서 고정
        )
        fig_monthly.update_traces(
            textposition="outside",
            textfont_size=11,
            marker_line_width=0,
            hovertemplate="<b>%{x}</b><br>매출액: %{customdata}<extra></extra>",
            customdata=monthly_sales["매출액_표시"],
        )
        fig_monthly.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            showlegend=False,
            coloraxis_showscale=False,
            height=400,
            margin=dict(t=20, b=40, l=40, r=20),
            xaxis=dict(
                type="category",           # datetime 오인 방지
                tickangle=-30,
                showgrid=False,
            ),
            yaxis=dict(
                showgrid=True,
                gridcolor="#f0f0f0",
                tickformat=",",
                tickprefix="₩",
            ),
            font=dict(family="sans-serif"),
        )
        st.plotly_chart(fig_monthly, width="stretch")

        # ── 차트 2 & 3: 2단 레이아웃 ──────────────────────────────────────
        col_left, col_right = st.columns([1, 1], gap="medium")

        # ── 거래처별 전체 집계 (공통 데이터) ─────────────────────────────
        client_sales = (
            df_sales.groupby("거래처")["부가세포함"]
            .sum()
            .reset_index()
            .rename(columns={"부가세포함": "매출액"})
            .sort_values("매출액", ascending=False)
        )

        # ── 차트 2: 거래처별 매출 비중 + 드릴다운 ────────────────────────
        with col_left:
            st.markdown('<div class="section-header">🏢 거래처별 매출 순위</div>', unsafe_allow_html=True)

            # ── 뷰 전환 토글: 전체 거래처 ↔ 특정 거래처 내 교수별 ─────────
            all_clients_sorted = client_sales["거래처"].tolist()

            # 세션키: 선택된 거래처 (None = 전체 보기)
            if "selected_client" not in st.session_state:
                st.session_state["selected_client"] = None

            # 상단 컨트롤 행
            ctrl_col1, ctrl_col2 = st.columns([3, 1])
            with ctrl_col1:
                # 현재 뷰 상태 표시
                if st.session_state["selected_client"] is None:
                    st.markdown(
                        '<span style="font-size:13px; color:#5a6478; font-weight:600;">'
                        '📊 전체 거래처 매출 비중</span>',
                        unsafe_allow_html=True,
                    )
                else:
                    selected_c = st.session_state["selected_client"]
                    c_total = int(client_sales.loc[
                        client_sales["거래처"] == selected_c, "매출액"
                    ].values[0])
                    st.markdown(
                        f'<span style="font-size:13px; color:#4F8BF9; font-weight:700;">'
                        f'🔍 {selected_c} &nbsp;·&nbsp; {format_currency(c_total)}</span>',
                        unsafe_allow_html=True,
                    )
            with ctrl_col2:
                if st.session_state["selected_client"] is not None:
                    if st.button("◀ 전체 보기", key="back_to_all", width="stretch"):
                        st.session_state["selected_client"] = None
                        st.rerun()

            # ── [전체 보기] 파이 차트 ─────────────────────────────────────
            if st.session_state["selected_client"] is None:
                # 상위 8개 + 기타 합산
                if len(client_sales) > 8:
                    top8 = client_sales.head(8).copy()
                    others_sum = int(client_sales.iloc[8:]["매출액"].sum())
                    others_row = pd.DataFrame([{"거래처": "기타", "매출액": others_sum}])
                    pie_df = pd.concat([top8, others_row], ignore_index=True)
                else:
                    pie_df = client_sales.copy()

                pie_df["매출액_표시"] = pie_df["매출액"].apply(format_currency)

                fig_pie = px.pie(
                    pie_df,
                    names="거래처",
                    values="매출액",
                    hole=0.42,
                    color_discrete_sequence=px.colors.qualitative.Set3,
                )
                fig_pie.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate=(
                        "<b>%{label}</b><br>"
                        "매출액: %{customdata}<br>"
                        "비율: %{percent}<extra></extra>"
                    ),
                    customdata=pie_df["매출액_표시"],
                    textfont_size=11,
                    pull=[0.03] * len(pie_df),  # 각 조각 살짝 분리
                )
                fig_pie.update_layout(
                    height=340,
                    margin=dict(t=10, b=10, l=10, r=10),
                    showlegend=False,
                    paper_bgcolor="white",
                    font=dict(family="sans-serif"),
                )
                st.plotly_chart(fig_pie, width="stretch")

                # 순위 테이블 + 클릭 선택 드롭다운
                st.markdown(
                    '<span style="font-weight:700; font-size:14px; color:#1a1a2e;">'
                    '📋 거래처 매출 순위 &nbsp;—&nbsp; '
                    '<span style="font-size:12px; color:#5a6478; font-weight:400;">'
                    '거래처를 선택하면 교수별 내역을 확인할 수 있습니다</span></span>',
                    unsafe_allow_html=True,
                )

                # 순위 테이블 (포맷팅)
                rank_df = client_sales.copy().reset_index(drop=True)
                rank_df.index = range(1, len(rank_df) + 1)
                total_rev = rank_df["매출액"].sum()
                rank_df["비율"] = (rank_df["매출액"] / total_rev * 100).round(1).astype(str) + "%"
                rank_df["매출액"] = rank_df["매출액"].apply(format_currency)
                st.dataframe(rank_df, width="stretch", height=180)

                # 거래처 선택 드롭다운
                st.markdown("<br>", unsafe_allow_html=True)
                drill_options = ["— 거래처를 선택하세요 —"] + all_clients_sorted
                drill_choice = st.selectbox(
                    "🔍 교수별 드릴다운",
                    options=drill_options,
                    key="drill_selectbox",
                    label_visibility="collapsed",
                )
                if drill_choice != "— 거래처를 선택하세요 —":
                    st.session_state["selected_client"] = drill_choice
                    st.rerun()

            # ── [드릴다운 뷰] 선택 거래처 내 교수별 매출 ────────────────────
            else:
                selected_c = st.session_state["selected_client"]

                # 해당 거래처 + 교수 필터 (교수 없는 행은 '미지정'으로 처리)
                drill_df = df_sales[df_sales["거래처"] == selected_c].copy()
                drill_df["교수"] = drill_df["교수"].replace("", "미지정")

                prof_drill = (
                    drill_df.groupby("교수")["부가세포함"]
                    .sum()
                    .reset_index()
                    .rename(columns={"부가세포함": "매출액"})
                    .sort_values("매출액", ascending=False)
                )
                prof_drill["매출액_표시"] = prof_drill["매출액"].apply(format_currency)
                drill_total = int(prof_drill["매출액"].sum())

                # 교수별 도넛 차트
                fig_drill = px.pie(
                    prof_drill,
                    names="교수",
                    values="매출액",
                    hole=0.45,
                    color_discrete_sequence=px.colors.qualitative.Pastel,
                )
                fig_drill.update_traces(
                    textposition="inside",
                    textinfo="percent+label",
                    hovertemplate=(
                        "<b>%{label}</b><br>"
                        "매출액: %{customdata}<br>"
                        "비율: %{percent}<extra></extra>"
                    ),
                    customdata=prof_drill["매출액_표시"],
                    textfont_size=11,
                    pull=[0.04] * len(prof_drill),
                )
                fig_drill.update_layout(
                    height=340,
                    margin=dict(t=10, b=10, l=10, r=10),
                    showlegend=False,
                    paper_bgcolor="white",
                    font=dict(family="sans-serif"),
                )
                st.plotly_chart(fig_drill, width="stretch")

                # 교수별 순위 테이블
                st.markdown(
                    f'<span style="font-weight:700; font-size:14px; color:#1a1a2e;">'
                    f'📋 {selected_c} · 교수별 매출 순위</span>',
                    unsafe_allow_html=True,
                )
                rank_drill = prof_drill.copy().reset_index(drop=True)
                rank_drill.index = range(1, len(rank_drill) + 1)
                rank_drill["비율"] = (
                    (rank_drill["매출액"] / drill_total * 100).round(1).astype(str) + "%"
                )
                rank_drill["매출액"] = rank_drill["매출액"].apply(format_currency)
                st.dataframe(rank_drill, width="stretch", height=180)

                # 거래건수 보조 정보
                cnt = len(drill_df)
                st.caption(f"총 {cnt:,}건 납품  |  교수 {len(prof_drill):,}명")

        # ── 차트 3: 교수별 매출 TOP 10 (전체 or 드릴다운 연동) ────────────
        with col_right:
            # 드릴다운 거래처가 선택돼 있으면 해당 거래처만, 아니면 전체
            selected_c = st.session_state.get("selected_client", None)

            if selected_c is None:
                st.markdown(
                    '<div class="section-header">👨‍🔬 교수별 매출 TOP 10 (전체)</div>',
                    unsafe_allow_html=True,
                )
                prof_base = df_sales[df_sales["교수"] != ""]
                chart_title_suffix = "전체"
            else:
                st.markdown(
                    f'<div class="section-header">👨‍🔬 {selected_c} · 교수별 매출</div>',
                    unsafe_allow_html=True,
                )
                prof_base = df_sales[
                    (df_sales["거래처"] == selected_c) & (df_sales["교수"] != "")
                ]
                chart_title_suffix = selected_c

            prof_sales = (
                prof_base
                .groupby("교수")["부가세포함"]
                .sum()
                .reset_index()
                .rename(columns={"부가세포함": "매출액"})
                .sort_values("매출액", ascending=False)
                .head(10)
                .sort_values("매출액", ascending=True)  # 가로 막대용 역정렬
            )
            prof_sales["매출액_표시"] = prof_sales["매출액"].apply(format_currency)

            # 색상: 드릴다운이면 파스텔 계열로 구분
            color_scale = (
                ["#c8e6f9", "#4F8BF9", "#1a4fa0"]
                if selected_c is None
                else ["#ffd8b1", "#ff9f43", "#e17055"]
            )

            fig_prof = px.bar(
                prof_sales,
                x="매출액",
                y="교수",
                orientation="h",
                text="매출액_표시",
                color="매출액",
                color_continuous_scale=color_scale,
                labels={"매출액": "매출액 (원)", "교수": ""},
            )
            fig_prof.update_traces(
                textposition="outside",
                textfont_size=10,
                marker_line_width=0,
                hovertemplate="<b>%{y}</b><br>매출액: %{customdata}<extra></extra>",
                customdata=prof_sales["매출액_표시"],
            )

            # 교수 수에 따라 차트 높이 동적 조절
            bar_height = max(320, min(500, len(prof_sales) * 42 + 60))
            fig_prof.update_layout(
                plot_bgcolor="white",
                paper_bgcolor="white",
                showlegend=False,
                coloraxis_showscale=False,
                height=bar_height,
                margin=dict(t=20, b=20, l=10, r=90),
                xaxis=dict(
                    showgrid=True,
                    gridcolor="#f0f0f0",
                    tickformat=",",
                    tickprefix="₩",
                ),
                yaxis=dict(showgrid=False, tickfont=dict(size=12)),
                font=dict(family="sans-serif"),
            )
            st.plotly_chart(fig_prof, width="stretch")

            # 드릴다운 시 추가 정보 표시
            if selected_c is not None and not prof_sales.empty:
                top_prof = prof_sales.sort_values("매출액", ascending=False).iloc[0]
                st.info(
                    f"🏆 **{selected_c}** 최고 매출 교수: "
                    f"**{top_prof['교수']}** · {format_currency(top_prof['매출액'])}"
                )

        # ── 교수별 월별 매출 추이 & 비교 ──────────────────────────────────
        st.markdown("---")
        st.markdown('<div class="section-header">📈 교수별 월별 매출 추이 & 비교</div>', unsafe_allow_html=True)

        # ── ① 전처리를 @st.cache_data 로 캐싱 — 교수 목록/피벗 사전 계산 ──
        @st.cache_data(show_spinner=False)
        def build_prof_month_pivot(sales_hash: int) -> pd.DataFrame:
            """
            교수 × 월 피벗을 미리 계산해 캐시.
            sales_hash 는 df_sales 변경 감지용 (len × sum 조합).
            반환: index=교수, columns=월(str), values=부가세포함 합계(int)
            """
            base = (
                df_sales[df_sales["교수"] != ""]
                .groupby(["교수", "월"])["부가세포함"]
                .sum()
                .reset_index()
                .rename(columns={"부가세포함": "매출액"})
            )
            # 모든 월 컬럼이 존재하도록 피벗
            pivot = base.pivot_table(
                index="교수", columns="월",
                values="매출액", aggfunc="sum", fill_value=0
            )
            pivot.columns = [str(c) for c in pivot.columns]  # 문자열 보장
            return pivot

        # 캐시 키: 행 수 + 부가세포함 합계 (데이터 변경 감지)
        _cache_key = len(df_sales) * 10000 + int(df_sales["부가세포함"].sum() % 1e9)
        prof_month_pivot = build_prof_month_pivot(_cache_key)

        all_profs_list = sorted(prof_month_pivot.index.tolist())
        all_months_list = sorted(prof_month_pivot.columns.tolist())  # 'YYYY-MM' 문자열

        # 드릴다운 거래처가 선택돼 있으면 해당 거래처 교수만 기본 후보로
        selected_c_now = st.session_state.get("selected_client", None)
        if selected_c_now:
            default_pool = sorted(
                df_sales[
                    (df_sales["거래처"] == selected_c_now) & (df_sales["교수"] != "")
                ]["교수"].unique().tolist()
            )
        else:
            default_pool = all_profs_list

        ctrl_l, ctrl_r = st.columns([3, 1])
        with ctrl_l:
            # 기본값: 캐시된 피벗에서 합계 상위 3명
            top3_default = (
                prof_month_pivot
                .loc[prof_month_pivot.index.isin(default_pool)]
                .sum(axis=1).nlargest(3).index.tolist()
                if default_pool else []
            )
            compare_profs = st.multiselect(
                "📌 비교할 교수를 선택하세요 (복수 선택 가능)",
                options=all_profs_list,
                default=top3_default,
                key="compare_profs_select",
            )
        with ctrl_r:
            chart_mode = st.radio(
                "차트 유형",
                options=["개별 선", "누적 막대"],
                index=0,
                horizontal=True,
                key="trend_chart_mode",
            )

        if not compare_profs:
            st.info("비교할 교수를 1명 이상 선택해주세요.")
        else:
            # ── ② 피벗에서 선택 교수 행만 슬라이싱 → melt → O(n) 연산만 ──
            sub_pivot = prof_month_pivot.loc[
                prof_month_pivot.index.isin(compare_profs), all_months_list
            ].copy()

            # 빠진 교수(pivot에 없는 경우) 0행 추가
            for p in compare_profs:
                if p not in sub_pivot.index:
                    sub_pivot.loc[p] = 0

            trend_df = (
                sub_pivot
                .reset_index()
                .melt(id_vars="교수", var_name="월", value_name="매출액")
                .sort_values(["월", "교수"])
            )
            # 매출액_표시: 정수 포맷 (0은 빈 문자열로 처리해 차트 레이블 깔끔하게)
            trend_df["매출액_표시"] = trend_df["매출액"].apply(
                lambda v: format_currency(v) if v > 0 else ""
            )

            palette = px.colors.qualitative.Bold

            # X축 타입을 명시적으로 'category'로 지정 → Plotly datetime 오인 방지
            xaxis_cfg = dict(
                type="category",      # ← 핵심: 문자열 카테고리로 강제
                tickangle=-30,
                showgrid=False,
                categoryorder="category ascending",
            )

            if chart_mode == "개별 선":
                fig_trend = px.line(
                    trend_df,
                    x="월", y="매출액", color="교수",
                    markers=True,
                    color_discrete_sequence=palette,
                    labels={"매출액": "매출액 (원)", "월": ""},
                    custom_data=["교수", "매출액_표시"],
                    category_orders={"월": all_months_list},
                )
                fig_trend.update_traces(
                    line_width=2.5, marker_size=7,
                    hovertemplate=(
                        "<b>%{customdata[0]}</b><br>"
                        "%{x}<br>매출액: %{customdata[1]}<extra></extra>"
                    ),
                )
            else:
                fig_trend = px.bar(
                    trend_df,
                    x="월", y="매출액", color="교수",
                    barmode="stack",
                    color_discrete_sequence=palette,
                    labels={"매출액": "매출액 (원)", "월": ""},
                    custom_data=["교수", "매출액_표시"],
                    category_orders={"월": all_months_list},
                )
                fig_trend.update_traces(
                    hovertemplate=(
                        "<b>%{customdata[0]}</b><br>"
                        "%{x}<br>매출액: %{customdata[1]}<extra></extra>"
                    ),
                )

            fig_trend.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                height=380,
                margin=dict(t=20, b=40, l=40, r=20),
                xaxis=xaxis_cfg,
                yaxis=dict(
                    showgrid=True, gridcolor="#f0f0f0",
                    tickformat=",", tickprefix="₩",
                ),
                legend=dict(
                    orientation="h", yanchor="bottom",
                    y=1.02, xanchor="left", x=0,
                    font=dict(size=11),
                ),
                font=dict(family="sans-serif"),
            )
            st.plotly_chart(fig_trend, width="stretch")

            # 교수별 월별 피벗 테이블 (이미 계산된 sub_pivot 재사용)
            with st.expander("📋 교수별 월별 매출 상세 테이블", expanded=False):
                disp_pivot = sub_pivot.copy()
                disp_pivot["합계"] = disp_pivot.sum(axis=1)
                disp_pivot = disp_pivot.sort_values("합계", ascending=False)
                st.dataframe(
                    disp_pivot.map(format_currency),
                    width="stretch",
                )

        # ── 원본 데이터 미리보기 ──────────────────────────────────────────
        with st.expander("📋 납품 원본 데이터 미리보기", expanded=False):
            preview_df = df_sales.copy()
            # 금액 컬럼 포맷팅
            for col in ["단가", "금액", "부가세포함"]:
                preview_df[col] = preview_df[col].apply(format_currency)
            preview_df["납품일"] = preview_df["납품일"].dt.strftime("%Y-%m-%d")
            st.dataframe(preview_df, width="stretch", height=300)


# =============================================================================
# Tab 2: 납품 조회 (Inquiry) — 거래처·교수·월별 납품 내역 빠른 조회
# =============================================================================
with tab2:
    if not has_sales:
        st.markdown(
            """
            <div class="empty-state">
                <div class="icon">🔍</div>
                <h3>납품 내역을 업로드해주세요</h3>
                <p>납품 내역 파일을 업로드하면 거래처·교수·월별로 납품 내역을 조회할 수 있습니다.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="section-header">🔍 납품 내역 빠른 조회</div>', unsafe_allow_html=True)
        st.caption("거래처 담당자가 납품 품목·금액을 문의할 때 즉시 안내할 수 있도록 설계된 조회 화면입니다.")

        # ── 필터 행 ───────────────────────────────────────────────────────
        filt_c1, filt_c2, filt_c3, filt_c4 = st.columns([2, 2, 2, 1])

        # 거래처 필터
        client_opts = ["전체"] + sorted(df_sales["거래처"].replace("", pd.NA).dropna().unique().tolist())
        with filt_c1:
            sel_client = st.selectbox("🏢 거래처", client_opts, key="inq_client")

        # 교수 필터 (거래처 선택에 따라 동적 변화)
        if sel_client == "전체":
            prof_opts = ["전체"] + sorted(
                df_sales[df_sales["교수"] != ""]["교수"].unique().tolist()
            )
        else:
            prof_opts = ["전체"] + sorted(
                df_sales[(df_sales["거래처"] == sel_client) & (df_sales["교수"] != "")]
                ["교수"].unique().tolist()
            )
        with filt_c2:
            sel_prof = st.selectbox("👨‍🔬 교수", prof_opts, key="inq_prof")

        # ── 월 필터 선택 항목: YYYY-MM 형식 목록 ────────────────────────
        month_opts = ["전체"] + sorted(df_sales["월"].astype(str).unique().tolist())
        with filt_c3:
            sel_month = st.selectbox("📅 월 선택", month_opts, key="inq_month")

        with filt_c4:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🔄 초기화", key="inq_reset", use_container_width=True):
                for k in ["inq_client", "inq_prof", "inq_month"]:
                    if k in st.session_state:
                        del st.session_state[k]
                st.rerun()

        # ── 필터 적용 ─────────────────────────────────────────────────────
        # 거래처·교수 필터는 그대로 적용
        inq_df = df_sales.copy()
        inq_df["월"] = inq_df["월"].astype(str)          # 문자열 통일
        if sel_client != "전체":
            inq_df = inq_df[inq_df["거래처"] == sel_client]
        if sel_prof != "전체":
            inq_df = inq_df[inq_df["교수"] == sel_prof]

        # ── 월 필터: 특정 월(MM) 선택 시 "연도 무관 같은 달" 전체 포함 ──
        # sel_month_mm : "01"~"12" / sel_month == "전체"이면 None
        sel_month_mm = None
        if sel_month != "전체":
            sel_month_mm = sel_month.split("-")[1]        # "2024-03" → "03"
            inq_df = inq_df[inq_df["월"].str.endswith(f"-{sel_month_mm}")]

        # KPI용 df는 필터 결과 그대로 사용
        inq_total = int(inq_df["부가세포함"].sum())
        inq_cnt   = len(inq_df)

        k1, k2, k3 = st.columns(3)
        with k1:
            render_kpi_card("조회 건수",    f"{inq_cnt:,}건",             "필터 조건 기준", "#4F8BF9")
        with k2:
            render_kpi_card("조회 매출 합계", format_currency(inq_total), "부가세 포함",   "#00b894")
        with k3:
            avg_inq = inq_total // inq_cnt if inq_cnt > 0 else 0
            render_kpi_card("건당 평균 금액", format_currency(avg_inq),   "부가세 포함",   "#e17055")

        st.markdown("<br>", unsafe_allow_html=True)

        # ── 월별 납품 금액 요약 차트 ──────────────────────────────────────
        if not inq_df.empty:
            # ── 차트 헤더 (선택 달 표시) ─────────────────────────────────
            if sel_month_mm:
                chart_title = f"📅 연도별 {int(sel_month_mm)}월 납품 금액 비교"
                chart_caption = (
                    f"데이터가 존재하는 연도의 {int(sel_month_mm)}월 납품 금액을 비교합니다. "
                    f"해당 연도에 납품 내역이 없으면 '데이터 없음'으로 표시됩니다."
                )
            else:
                chart_title = "📅 월별 납품 금액 요약"
                chart_caption = ""

            st.markdown(f'<div class="section-header">{chart_title}</div>',
                        unsafe_allow_html=True)
            if chart_caption:
                st.caption(chart_caption)

            # ── X축 레이블 생성 로직 분기 ─────────────────────────────────
            if sel_month_mm:
                # ── [특정 달 선택] 연도별 비교 막대 ─────────────────────
                # 데이터에 있는 연도 추출
                years_in_data = sorted(inq_df["월"].str[:4].unique().tolist())

                # 5년 범위: 데이터 최소 연도~최대 연도 (최소 1개 보장)
                year_min = int(years_in_data[0])
                year_max = int(years_in_data[-1])
                all_years = [str(y) for y in range(year_min, year_max + 1)]

                # 연도별 매출액 집계 (없는 연도 = 0)
                year_sales = (
                    inq_df.groupby(inq_df["월"].str[:4])["부가세포함"]
                    .sum()
                    .reindex(all_years, fill_value=0)
                    .reset_index()
                    .rename(columns={"월": "연도", "부가세포함": "매출액"})
                )
                # X축 레이블: "YYYY년 MM월", 데이터 없으면 "YYYY년 (데이터 없음)"
                year_sales["X레이블"] = year_sales.apply(
                    lambda r: (
                        f"{r['연도']}년 {int(sel_month_mm)}월"
                        if r["매출액"] > 0
                        else f"{r['연도']}년\n(데이터 없음)"
                    ),
                    axis=1,
                )
                year_sales["매출액_표시"] = year_sales["매출액"].apply(
                    lambda v: format_currency(v) if v > 0 else "데이터 없음"
                )
                # 데이터 있음/없음 구분 색상
                year_sales["색상"] = year_sales["매출액"].apply(
                    lambda v: "데이터 있음" if v > 0 else "데이터 없음"
                )
                x_order = year_sales["X레이블"].tolist()

                fig_minq = px.bar(
                    year_sales,
                    x="X레이블", y="매출액",
                    text="매출액_표시",
                    color="색상",
                    color_discrete_map={
                        "데이터 있음": "#4F8BF9",
                        "데이터 없음": "#d0d7e6",
                    },
                    labels={"매출액": "매출액 (원)", "X레이블": ""},
                    category_orders={"X레이블": x_order},
                    custom_data=["매출액_표시", "연도"],
                )
                fig_minq.update_traces(
                    textposition="outside",
                    textfont_size=11,
                    marker_line_width=0,
                    hovertemplate=(
                        "<b>%{customdata[1]}년 "
                        + f"{int(sel_month_mm)}월</b><br>"
                        + "%{customdata[0]}<extra></extra>"
                    ),
                )
                fig_minq.update_layout(showlegend=True)

            else:
                # ── [전체 선택] 기존 월별 집계 (거래처/교수 필터만 적용) ─
                monthly_inq = (
                    inq_df.groupby("월")["부가세포함"]
                    .sum().reset_index()
                    .rename(columns={"부가세포함": "매출액"})
                    .sort_values("월")
                )
                monthly_inq["매출액_표시"] = monthly_inq["매출액"].apply(format_currency)
                inq_months_sorted = monthly_inq["월"].tolist()

                has_multi_prof = (
                    inq_df[inq_df["교수"] != ""]["교수"].nunique() > 1
                    and sel_prof == "전체"
                )

                if has_multi_prof:
                    monthly_inq_prof = (
                        inq_df[inq_df["교수"] != ""]
                        .groupby(["월", "교수"])["부가세포함"]
                        .sum().reset_index()
                        .rename(columns={"부가세포함": "매출액"})
                        .sort_values("월")
                    )
                    monthly_inq_prof["매출액_표시"] = \
                        monthly_inq_prof["매출액"].apply(format_currency)
                    fig_minq = px.bar(
                        monthly_inq_prof,
                        x="월", y="매출액", color="교수",
                        barmode="stack",
                        color_discrete_sequence=px.colors.qualitative.Pastel,
                        labels={"매출액": "매출액 (원)", "월": ""},
                        custom_data=["교수", "매출액_표시"],
                        category_orders={"월": inq_months_sorted},
                    )
                    fig_minq.update_traces(
                        hovertemplate=(
                            "<b>%{customdata[0]}</b><br>"
                            "%{x}<br>%{customdata[1]}<extra></extra>"
                        ),
                    )
                else:
                    fig_minq = px.bar(
                        monthly_inq,
                        x="월", y="매출액", text="매출액_표시",
                        color="매출액",
                        color_continuous_scale=["#a8d8ea", "#4F8BF9", "#1a1a9e"],
                        labels={"매출액": "매출액 (원)", "월": ""},
                        category_orders={"월": inq_months_sorted},
                    )
                    fig_minq.update_traces(
                        textposition="outside", textfont_size=11,
                        hovertemplate="<b>%{x}</b><br>%{customdata}<extra></extra>",
                        customdata=monthly_inq["매출액_표시"],
                    )

            # ── 공통 레이아웃 ────────────────────────────────────────────
            n_bars = len(year_sales) if sel_month_mm else len(
                inq_df["월"].unique()
            )
            chart_h = max(300, min(420, n_bars * 55 + 80))
            fig_minq.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                coloraxis_showscale=False,
                height=chart_h,
                margin=dict(t=20, b=50, l=40, r=20),
                xaxis=dict(
                    type="category",
                    categoryorder="array",
                    tickangle=-20,
                    showgrid=False,
                    tickfont=dict(size=12),
                ),
                yaxis=dict(
                    showgrid=True, gridcolor="#f0f0f0",
                    tickformat=",", tickprefix="₩",
                ),
                legend=dict(
                    orientation="h", yanchor="bottom", y=1.02,
                    xanchor="left", x=0, font=dict(size=11),
                ),
                font=dict(family="sans-serif"),
            )
            st.plotly_chart(fig_minq, width="stretch")

            # ── 교수별 납품 금액 피벗 (월별) ─────────────────────────────
            st.markdown('<div class="section-header">👨‍🔬 교수별 · 월별 납품 금액</div>', unsafe_allow_html=True)

            pivot_inq = (
                inq_df[inq_df["교수"] != ""]
                .groupby(["교수", "월"])["부가세포함"]
                .sum()
                .reset_index()
                .pivot_table(index="교수", columns="월",
                             values="부가세포함", aggfunc="sum", fill_value=0)
            )
            if not pivot_inq.empty:
                pivot_inq["합계"] = pivot_inq.sum(axis=1)
                pivot_inq = pivot_inq.sort_values("합계", ascending=False)
                pivot_inq_fmt = pivot_inq.map(format_currency)
                st.dataframe(pivot_inq_fmt, width="stretch")
            else:
                st.info("교수 정보가 없습니다.")

            # ── 상세 납품 내역 테이블 ─────────────────────────────────────
            st.markdown('<div class="section-header">📋 상세 납품 내역</div>', unsafe_allow_html=True)

            # 품목별 요약 (동일 거래처 문의 대응용)
            tab_sum, tab_detail = st.tabs(["📦 품목별 요약", "📄 건별 상세"])

            with tab_sum:
                summary_inq = (
                    inq_df.groupby(["Brand", "Cat.No", "품명", "사이즈"])
                    .agg(
                        총수량=("수량", "sum"),
                        총금액=("금액", "sum"),
                        부가세포함=("부가세포함", "sum"),
                        납품건수=("납품일", "count"),
                    )
                    .reset_index()
                    .sort_values("부가세포함", ascending=False)
                )
                summary_inq["총금액"]   = summary_inq["총금액"].apply(format_currency)
                summary_inq["부가세포함"] = summary_inq["부가세포함"].apply(format_currency)
                summary_inq = summary_inq.reset_index(drop=True)
                summary_inq.index = range(1, len(summary_inq) + 1)
                st.dataframe(summary_inq, width="stretch", height=300)
                st.caption(f"총 {len(summary_inq):,}개 품목 | 합계 {format_currency(inq_total)}")

            with tab_detail:
                detail_inq = inq_df[[
                    "납품일", "거래처", "교수", "연구원",
                    "Brand", "Cat.No", "품명", "사이즈",
                    "수량", "단가", "금액", "부가세포함",
                ]].copy().sort_values("납품일", ascending=False)
                detail_inq["납품일"] = detail_inq["납품일"].dt.strftime("%Y-%m-%d")
                for col in ["단가", "금액", "부가세포함"]:
                    detail_inq[col] = detail_inq[col].apply(format_currency)
                detail_inq = detail_inq.reset_index(drop=True)
                detail_inq.index = range(1, len(detail_inq) + 1)
                st.dataframe(detail_inq, width="stretch", height=350)
                st.caption(f"총 {len(detail_inq):,}건")

                # 엑셀 다운로드
                @st.cache_data
                def to_excel_bytes(df_raw: pd.DataFrame) -> bytes:
                    out = BytesIO()
                    export_df = df_raw.copy()
                    export_df["납품일"] = pd.to_datetime(export_df["납품일"], errors="coerce")
                    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                        export_df.to_excel(writer, index=False, sheet_name="납품내역")
                    return out.getvalue()

                raw_export = inq_df[[
                    "납품일", "거래처", "교수", "연구원",
                    "Brand", "Cat.No", "품명", "사이즈",
                    "수량", "단가", "금액", "부가세포함",
                ]].copy().sort_values("납품일", ascending=False)

                fname_parts = []
                if sel_client != "전체": fname_parts.append(sel_client)
                if sel_prof   != "전체": fname_parts.append(sel_prof)
                if sel_month  != "전체": fname_parts.append(sel_month)
                fname = ("_".join(fname_parts) if fname_parts else "전체") + "_납품내역.xlsx"

                st.download_button(
                    label="⬇️ 조회 결과 엑셀 다운로드",
                    data=to_excel_bytes(raw_export),
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.warning("선택한 조건에 해당하는 납품 내역이 없습니다.")


# =============================================================================
# Tab 3: 장부 관리 (Ledger)
# =============================================================================
with tab3:
    if not has_sales and not has_payment:
        # 업로드 안내
        st.markdown(
            """
            <div class="empty-state">
                <div class="icon">📒</div>
                <h3>파일을 업로드해주세요</h3>
                <p>납품 내역과 결제 내역 파일을 모두 업로드하면<br>
                교수별 미수금을 확인할 수 있습니다.</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        # ── 교수 선택 드롭다운 ────────────────────────────────────────────
        st.markdown('<div class="section-header">👨‍🔬 교수별 미수금 관리</div>', unsafe_allow_html=True)

        # 납품 내역 & 결제 내역에 등장하는 교수 목록 통합 (빈 값 제외)
        professors_from_sales = (
            set(df_sales[df_sales["교수"] != ""]["교수"].unique())
            if has_sales else set()
        )
        professors_from_payment = (
            set(df_payment[df_payment["교수"] != ""]["교수"].unique())
            if has_payment else set()
        )
        all_professors = sorted(professors_from_sales | professors_from_payment)

        if not all_professors:
            st.warning("교수 데이터가 없습니다. 파일의 '교수' 컬럼을 확인해주세요.")
        else:
            # 교수별 미수금 요약 계산 (드롭다운 옵션용)
            def get_receivable(prof: str) -> int:
                """교수 이름으로 미수금(총 구매액 - 총 결제액) 계산"""
                sales_total = (
                    int(df_sales[df_sales["교수"] == prof]["부가세포함"].sum())
                    if has_sales else 0
                )
                payment_total = (
                    int(df_payment[df_payment["교수"] == prof]["금액"].sum())
                    if has_payment else 0
                )
                return sales_total - payment_total

            # 드롭다운 레이블에 미수금 정보 포함
            professor_options = []
            for prof in all_professors:
                recv = get_receivable(prof)
                badge = "🔴" if recv > 0 else "🟢"
                professor_options.append(f"{badge} {prof}  ({format_currency(recv)})")

            col_sel, col_summary = st.columns([1, 2])
            with col_sel:
                selected_option = st.selectbox(
                    "교수를 선택하세요",
                    options=professor_options,
                    index=0,
                    label_visibility="collapsed",
                )

            # 선택된 교수 이름 파싱 (이모지, 금액 제거)
            selected_prof = selected_option.split("  ")[0].replace("🔴 ", "").replace("🟢 ", "").strip()

            # ── 해당 교수 KPI 계산 ────────────────────────────────────────
            sales_total = (
                int(df_sales[df_sales["교수"] == selected_prof]["부가세포함"].sum())
                if has_sales else 0
            )
            payment_total = (
                int(df_payment[df_payment["교수"] == selected_prof]["금액"].sum())
                if has_payment else 0
            )
            receivable = sales_total - payment_total

            st.markdown("<br>", unsafe_allow_html=True)

            # ── 미수금 카드 ───────────────────────────────────────────────
            col_s, col_p, col_r = st.columns(3)

            with col_s:
                render_kpi_card(
                    "📦 총 구매액",
                    format_currency(sales_total),
                    "부가세 포함 합계",
                    "#4F8BF9",
                )
            with col_p:
                render_kpi_card(
                    "💳 총 결제액",
                    format_currency(payment_total),
                    "입금 완료 금액",
                    "#00b894",
                )
            with col_r:
                # 미수금 상태에 따라 색상 변경
                if receivable > 0:
                    card_class = "alert-danger"
                    status_text = f"⚠️ 미수금 {format_currency(receivable)} 남아있음"
                    color = "#ee5a24"
                elif receivable < 0:
                    card_class = "alert-success"
                    status_text = f"✅ 초과 결제 {format_currency(abs(receivable))}"
                    color = "#01a3a4"
                else:
                    card_class = "alert-success"
                    status_text = "✅ 결제 완료"
                    color = "#01a3a4"

                render_kpi_card(
                    "🧾 미수금 (구매액 - 결제액)",
                    format_currency(receivable),
                    status_text,
                    color,
                )

            st.markdown("---")

            # ── 미수금 상태 배너 ──────────────────────────────────────────
            if receivable > 0:
                st.markdown(
                    f"""
                    <div class="alert-danger">
                        <div style="font-size:18px; font-weight:700; margin-bottom:4px;">
                            ⚠️ 미수금이 존재합니다
                        </div>
                        <div class="alert-value">{format_currency(receivable)}</div>
                        <div class="alert-label">{selected_prof} 교수 · 추가 결제 요청 필요</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    f"""
                    <div class="alert-success">
                        <div style="font-size:18px; font-weight:700; margin-bottom:4px;">
                            ✅ 결제 완료
                        </div>
                        <div class="alert-value">{format_currency(receivable)}</div>
                        <div class="alert-label">{selected_prof} 교수 · 미수금 없음</div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)

            # ── 상세 내역 테이블 2단 레이아웃 ──────────────────────────────
            col_detail_l, col_detail_r = st.columns([3, 2], gap="large")

            # ── 납품(구매) 상세 내역 ──────────────────────────────────────
            with col_detail_l:
                st.markdown(f"**📦 납품(구매) 상세 내역** - {selected_prof} 교수")

                if has_sales:
                    prof_sales_df = df_sales[df_sales["교수"] == selected_prof].copy()

                    if prof_sales_df.empty:
                        st.info("납품 내역이 없습니다.")
                    else:
                        # 표시용 포맷팅
                        display_sales = prof_sales_df[
                            ["납품일", "거래처", "연구원", "Brand", "Cat.No", "품명",
                             "사이즈", "수량", "단가", "금액", "부가세포함"]
                        ].copy()
                        display_sales["납품일"] = display_sales["납품일"].dt.strftime("%Y-%m-%d")
                        for col in ["단가", "금액", "부가세포함"]:
                            display_sales[col] = display_sales[col].apply(format_currency)
                        display_sales = display_sales.reset_index(drop=True)
                        display_sales.index = range(1, len(display_sales) + 1)

                        st.dataframe(
                            display_sales,
                            width="stretch",
                            height=350,
                        )
                        st.caption(f"총 {len(prof_sales_df):,}건 | 합계 {format_currency(sales_total)}")
                else:
                    st.info("납품 내역 파일을 업로드해주세요.")

            # ── 결제 상세 내역 ────────────────────────────────────────────
            with col_detail_r:
                st.markdown(f"**💳 결제 상세 내역** - {selected_prof} 교수")

                if has_payment:
                    prof_payment_df = df_payment[df_payment["교수"] == selected_prof].copy()

                    if prof_payment_df.empty:
                        st.info("결제 내역이 없습니다.")
                    else:
                        display_payment = prof_payment_df[
                            ["결제일", "교수", "연구원", "금액"]
                        ].copy()
                        display_payment["결제일"] = display_payment["결제일"].dt.strftime("%Y-%m-%d")
                        display_payment["금액"] = display_payment["금액"].apply(format_currency)
                        display_payment = display_payment.reset_index(drop=True)
                        display_payment.index = range(1, len(display_payment) + 1)

                        st.dataframe(
                            display_payment,
                            width="stretch",
                            height=350,
                        )
                        st.caption(f"총 {len(prof_payment_df):,}건 | 합계 {format_currency(payment_total)}")
                else:
                    st.info("결제 내역 파일을 업로드해주세요.")

            st.markdown("---")

            # ── 전체 교수 미수금 요약 테이블 ──────────────────────────────
            st.markdown('<div class="section-header">📋 전체 교수 미수금 요약</div>', unsafe_allow_html=True)

            # 전체 교수별 구매/결제/미수금 집계
            summary_rows = []
            for prof in all_professors:
                s_total = (
                    int(df_sales[df_sales["교수"] == prof]["부가세포함"].sum())
                    if has_sales else 0
                )
                p_total = (
                    int(df_payment[df_payment["교수"] == prof]["금액"].sum())
                    if has_payment else 0
                )
                recv = s_total - p_total
                summary_rows.append({
                    "교수": prof,
                    "총 구매액": s_total,
                    "총 결제액": p_total,
                    "미수금": recv,
                    "상태": "⚠️ 미수금" if recv > 0 else ("✅ 완료" if recv == 0 else "🔵 초과"),
                })

            summary_df = pd.DataFrame(summary_rows).sort_values("미수금", ascending=False)

            # 금액 포맷팅 (표시용 복사본)
            display_summary = summary_df.copy()
            for col in ["총 구매액", "총 결제액", "미수금"]:
                display_summary[col] = display_summary[col].apply(format_currency)
            display_summary = display_summary.reset_index(drop=True)
            display_summary.index = range(1, len(display_summary) + 1)

            # 미수금 있는 행 강조를 위한 스타일 함수
            def highlight_receivable(row):
                """미수금 > 0인 행을 연한 붉은색으로 강조"""
                raw_recv = summary_df.loc[
                    summary_df["교수"] == row["교수"], "미수금"
                ].values
                if len(raw_recv) > 0 and raw_recv[0] > 0:
                    return ["background-color: #fff5f5"] * len(row)
                return [""] * len(row)

            st.dataframe(
                display_summary,
                width="stretch",
                height=min(400, 35 * len(display_summary) + 40),
            )

            # 전체 미수금 합계
            total_receivable = summary_df["미수금"].sum()
            col_tr1, col_tr2, col_tr3 = st.columns([1, 1, 1])
            with col_tr1:
                st.metric("전체 교수 수", f"{len(summary_df)}명")
            with col_tr2:
                outstanding_count = len(summary_df[summary_df["미수금"] > 0])
                st.metric("미수금 교수 수", f"{outstanding_count}명",
                          delta=f"-{outstanding_count}명 관리 필요" if outstanding_count > 0 else "0",
                          delta_color="inverse")
            with col_tr3:
                st.metric(
                    "전체 미수금 합계",
                    format_currency(total_receivable),
                    delta=None,
                )


# =============================================================================
# 8. 샘플 데이터 다운로드 (사이드바 하단 추가)
# =============================================================================
with st.sidebar:
    st.markdown("---")
    st.markdown("### 📥 샘플 데이터")

    @st.cache_data
    def create_sample_excel():
        """테스트용 샘플 엑셀 파일 생성 (메모리 내)"""
        import random
        from datetime import datetime, timedelta

        # 샘플 납품 내역 생성
        brands = ["Sigma-Aldrich", "Thermo Fisher", "Merck", "Bio-Rad", "Abcam"]
        cat_nos = ["A1234", "B5678", "C9012", "D3456", "E7890", "F1122", "G3344"]
        products = ["Trypsin", "DMEM", "PBS Buffer", "EDTA", "BSA", "Fetal Bovine Serum",
                    "Antibody Anti-CD3", "PCR Master Mix", "Agarose", "Ethidium Bromide"]
        sizes = ["100ml", "500ml", "1L", "1mg", "5mg", "25mg", "100ug"]
        clients = ["서울대학교병원", "연세의료원", "삼성서울병원", "서울아산병원", "고려대안암병원"]
        professors = ["김민준 교수", "이서연 교수", "박지훈 교수", "최수빈 교수", "정예린 교수",
                      "강현우 교수", "윤지아 교수", "임도현 교수"]
        researchers = ["김연구원", "이연구원", "박연구원", "최연구원", "정연구원"]

        sales_rows = []
        base_date = datetime(2024, 1, 1)
        for i in range(120):
            delivery_date = base_date + timedelta(days=random.randint(0, 365))
            client = random.choice(clients)
            prof = random.choice(professors)
            researcher = random.choice(researchers)
            brand = random.choice(brands)
            cat_no = random.choice(cat_nos)
            product = random.choice(products)
            size = random.choice(sizes)
            qty = random.randint(1, 10)
            unit_price = random.choice([50000, 80000, 120000, 150000, 200000, 350000, 500000])
            amount = qty * unit_price
            vat_amount = int(amount * 1.1)

            sales_rows.append([
                delivery_date, client, prof, researcher, brand, cat_no,
                product, size, qty, unit_price, amount, vat_amount
            ])

        sales_df = pd.DataFrame(sales_rows, columns=[
            "납품일", "거래처", "교수", "연구원", "Brand", "Cat.No",
            "품명", "사이즈", "수량", "단가", "금액", "부가세포함"
        ])

        # 샘플 결제 내역 생성
        payment_rows = []
        for prof in professors:
            # 교수당 2~5건 결제
            for _ in range(random.randint(2, 5)):
                pay_date = base_date + timedelta(days=random.randint(30, 365))
                researcher = random.choice(researchers)
                pay_amount = random.choice([200000, 500000, 800000, 1000000, 1500000])
                payment_rows.append([pay_date, prof, researcher, pay_amount])

        payment_df = pd.DataFrame(payment_rows, columns=["결제일", "교수", "연구원", "금액"])

        # BytesIO에 두 시트로 저장
        output_sales = BytesIO()
        output_payment = BytesIO()

        with pd.ExcelWriter(output_sales, engine="xlsxwriter") as writer:
            sales_df.to_excel(writer, index=False, sheet_name="납품내역")

        with pd.ExcelWriter(output_payment, engine="xlsxwriter") as writer:
            payment_df.to_excel(writer, index=False, sheet_name="결제내역")

        return output_sales.getvalue(), output_payment.getvalue()

    # 샘플 파일 생성
    sample_sales_bytes, sample_payment_bytes = create_sample_excel()

    st.download_button(
        label="📦 납품 내역 샘플 다운로드",
        data=sample_sales_bytes,
        file_name="샘플_납품내역.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
    st.download_button(
        label="💳 결제 내역 샘플 다운로드",
        data=sample_payment_bytes,
        file_name="샘플_결제내역.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch",
    )
