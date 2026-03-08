# 🧬 생명공학 시약 영업 관리 ERP

생명공학 시약 기술영업 담당자를 위한 개인용 ERP 시스템입니다.  
납품 내역과 결제 내역 엑셀 파일을 업로드하면 매출 현황을 시각화하고 교수별 미수금을 관리할 수 있습니다.

---

## 🚀 배포 URL

| 환경 | URL |
|---|---|
| **Streamlit Cloud (영구)** | https://sungbinerp.streamlit.app *(배포 후 확정)* |
| **GitHub** | https://github.com/sungbin3346/-sungbinerp |

---

## ✅ 주요 기능

### 📊 대시보드
- **KPI 카드**: 총 매출액, 총 거래 건수, 거래처 수, 담당 교수 수
- **월별 매출 추이** 막대 차트
- **거래처별 매출 비중** 도넛 파이 차트 + 교수별 드릴다운
- **교수별 매출 TOP 10** 가로 막대 차트
- **교수별 월별 추이 & 비교** (멀티셀렉트, 선형/누적막대)

### 🔍 납품 조회
- 거래처 · 교수 · 월 3단 필터
- 월별 납품 금액 요약 차트 (특정 달 선택 시 연도별 비교 모드)
- 교수별×월별 피벗 테이블
- 품목별 요약 / 건별 상세 테이블
- 엑셀 다운로드

### 📒 장부 관리
- 교수별 미수금 = 총 구매액 − 총 결제액
- 색상 배지 (🔴 미수금 / 🟢 완료)
- 납품 · 결제 상세 내역 테이블
- 전체 교수 미수금 요약 테이블

---

## 📦 데이터 구조

### 납품 내역 (납품 내역.xlsx)
| 컬럼 | 설명 |
|---|---|
| 납품일 | 날짜 (YYYY-MM-DD) |
| 거래처 | 병원/기관명 |
| 교수 | 담당 교수명 |
| 연구원 | 담당 연구원명 |
| Brand | 제품 브랜드 |
| Cat.No | 카탈로그 번호 |
| 품명 | 제품명 |
| 사이즈 | 제품 사이즈 |
| 수량 | 수량 |
| 단가 | 단가 (원) |
| 금액 | 금액 (원) |
| 부가세포함 | 부가세 포함 금액 (원) |

### 결제 내역 (결제 내역.xlsx)
| 컬럼 | 설명 |
|---|---|
| 결제일 | 날짜 (YYYY-MM-DD) |
| 교수 | 담당 교수명 |
| 연구원 | 담당 연구원명 |
| 금액 | 입금액 (원) |

---

## 💾 데이터 저장 (Google Sheets 영구 저장)

Streamlit Cloud 배포 시 **Google Sheets**를 통해 데이터를 영구 저장합니다.

### 설정 방법
1. [Google Cloud Console](https://console.cloud.google.com) → 새 프로젝트 생성
2. **Google Sheets API** + **Google Drive API** 활성화
3. 서비스 계정 생성 → JSON 키 다운로드
4. Streamlit Cloud Secrets에 아래 내용 입력:

```toml
# .streamlit/secrets.toml (로컬 테스트용 — git에 올리지 말 것)
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email = "your-service@your-project.iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"

SPREADSHEET_ID = "your-google-spreadsheet-id"
```

5. 서비스 계정 이메일을 Google Sheets에 **편집자**로 공유

---

## 🖥️ 로컬 실행

```bash
# 1. 의존성 설치
pip install -r requirements.txt

# 2. Google Sheets 비밀키 설정 (선택)
mkdir -p .streamlit
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# → secrets.toml 에 실제 값 입력

# 3. 앱 실행
streamlit run app.py
```

---

## 📱 iPhone 홈 화면 추가 (PWA)

1. iPhone Safari에서 앱 URL 접속
2. 하단 **공유 버튼** (□↑) 탭
3. **"홈 화면에 추가"** 선택
4. 이름 확인 후 **추가** 탭
5. 홈 화면에서 앱 아이콘으로 실행 🎉

---

## 🛠️ 기술 스택

- **Frontend/Backend**: Python + Streamlit
- **데이터 처리**: pandas
- **시각화**: Plotly Express / Graph Objects
- **영구 저장**: Google Sheets (gspread)
- **PWA**: manifest.json + Service Worker
- **배포**: Streamlit Community Cloud
