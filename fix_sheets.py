import re

with open("sheets_db.py", "r") as f:
    content = f.read()

# _get_spreadsheet 함수에서 SPREADSHEET_ID 읽는 부분 수정
old = 'spreadsheet_id = st.secrets.get("SPREADSHEET_ID", "")'
new = '''try:
            spreadsheet_id = st.secrets["SPREADSHEET_ID"]
        except Exception:
            try:
                spreadsheet_id = st.secrets["gcp_service_account"].get("spreadsheet_id", "")
            except Exception:
                spreadsheet_id = ""'''

content = content.replace(old, new)

# is_connected 함수도 수정
old2 = 'sid = st.secrets.get("SPREADSHEET_ID", "")'
new2 = '''try:
            sid = st.secrets["SPREADSHEET_ID"]
        except Exception:
            try:
                sid = st.secrets["gcp_service_account"].get("spreadsheet_id", "")
            except Exception:
                sid = ""'''

content = content.replace(old2, new2)

with open("sheets_db.py", "w") as f:
    f.write(content)

print("수정 완료!")
