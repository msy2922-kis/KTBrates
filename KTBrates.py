import streamlit as st
from email.message import EmailMessage
import datetime

# 1. 페이지 설정 (최소화)
st.set_page_config(page_title="Fast Rate Sender", layout="centered")

# 2. 캐싱된 파일 리더 함수 (Pandas 제거 -> 직접 읽기로 속도 극대화)
@st.cache_data(show_spinner=False)
def get_rates_from_excel(uploaded_file):
    """
    Pandas를 쓰지 않고 엑셀 라이브러리를 직접 사용하여 속도를 높입니다.
    E2(row1, col4), L2(row1, col11), P2(row1, col15) 값을 가져옵니다.
    """
    try:
        filename = uploaded_file.name.lower()
        
        # .xls 파일인 경우 (구버전 엑셀)
        if filename.endswith('.xls'):
            import xlrd
            wb = xlrd.open_workbook(file_contents=uploaded_file.read())
            sheet = wb.sheet_by_index(0)
            # xlrd는 (row, col) 순서, 0부터 시작
            v_3m = sheet.cell_value(1, 4)   # E2
            v_3y = sheet.cell_value(1, 11)  # L2
            v_10y = sheet.cell_value(1, 15) # P2
            
        # .xlsx 파일인 경우 (신버전 엑셀)
        else:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded_file, data_only=True, read_only=True)
            sheet = wb.active
            # openpyxl은 1부터 시작 (row=2, column=5는 E2)
            v_3m = sheet.cell(row=2, column=5).value
            v_3y = sheet.cell(row=2, column=12).value
            v_10y = sheet.cell(row=2, column=16).value

        # 소수점 처리 (혹시 숫자로 들어오면 문자로 변환)
        def fmt(val):
            return str(val) if val is not None else ""
            
        return fmt(v_3m), fmt(v_3y), fmt(v_10y)

    except Exception:
        return "", "", ""

# 3. 설정 로드
secrets = st.secrets.get("gmail", {})
sender_email = secrets.get("id", "")
sender_password = secrets.get("pw", "")
default_receiver = secrets.get("receiver", "")

# 4. UI 구성 (심플함 유지)
st.markdown("### ⚡ Fast Rate Sender")

# 파일 업로드 (가장 먼저)
uploaded_file = st.file_uploader("Excel Upload", type=["xls", "xlsx"], label_visibility="collapsed")

# 기본값 설정
v_3m, v_3y, v_10y = "", "", ""

# 파일이 있으면 즉시 분석 (캐싱됨)
if uploaded_file:
    v_3m, v_3y, v_10y = get_rates_from_excel(uploaded_file)

# 5. 입력 및 전송 폼 (Form)
with st.form("fast_form"):
    # UI 렌더링 속도를 위해 컬럼 없이 일자 배치
    receiver = st.text_input("To", value=default_receiver)
    cd = st.text_input("CD (%)", placeholder="Direct Input")
    
    # 엑셀에서 가져온 값이 있으면 채워넣기
    k3m = st.text_input("KTB 3M", value=v_3m)
    k3y = st.text_input("KTB 3Y", value=v_3y)
    k10y = st.text_input("KTB 10Y", value=v_10y)

    # 전송 버튼
    submitted = st.form_submit_button("Send Mail", type="primary")

# 6. 전송 로직 (SMTP는 버튼 누를 때만 import)
if submitted:
    if not (sender_email and sender_password and receiver):
        st.error("Check Secrets!")
    elif not (cd and k3m and k3y and k10y):
        st.warning("Input Data!")
    else:
        try:
            import smtplib # 여기서 import하여 초기 로딩 속도 향상
            
            msg = EmailMessage()
            # pandas.Timestamp 대신 가벼운 datetime 사용
            today = datetime.datetime.now().strftime('%Y-%m-%d')
            msg['Subject'] = f"[Report] Market Rates {today}"
            msg['From'] = sender_email
            msg['To'] = receiver
            
            body = f"""Rates Report:
- CD: {cd}%
- KTB 3M: {k3m}%
- KTB 3Y: {k3y}%
- KTB 10Y: {k10y}%
"""
            msg.set_content(body)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender_email, sender_password)
                smtp.send_message(msg)
            
            st.success("Sent!")
            
        except Exception as e:
            st.error(f"Err: {e}")
