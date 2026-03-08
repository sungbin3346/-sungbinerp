import streamlit as st

st.title("Secrets 디버깅 테스트")

# 1) Secrets 전체 확인
try:
    all_keys = list(st.secrets.keys())
    st.success(f"Secrets 키 목록: {all_keys}")
except Exception as e:
    st.error(f"Secrets 자체를 읽을 수 없음: {e}")

# 2) gcp_service_account 확인
try:
    gcp = st.secrets["gcp_service_account"]
    st.success(f"gcp_service_account 읽기 성공! project_id = {gcp.get('project_id', 'N/A')}")
except Exception as e:
    st.error(f"gcp_service_account 읽기 실패: {e}")

# 3) SPREADSHEET_ID 확인
try:
    sid = st.secrets["SPREADSHEET_ID"]
    st.success(f"SPREADSHEET_ID = {sid}")
except Exception as e:
    st.error(f"SPREADSHEET_ID 읽기 실패: {e}")
