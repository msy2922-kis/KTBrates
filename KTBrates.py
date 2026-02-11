import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage

# 1. 페이지 설정: 제목 최소화, 레이아웃 'centered' 유지
st.set_page_config(page_title="Market Rate Sender", layout="centered")

# 2. 비밀번호/계정 로드 (UI 표시 안 함, 백그라운드 처리)
secrets = st.secrets.get("gmail", {})
sender_email = secrets.get("id", "")
sender_password = secrets.get("pw", "")
default_receiver = secrets.get("receiver", "")

# 3. 타이틀 최소화
st.header("Daily Rate Sender")

# 4. 파일 업로드 (가장 상단 배치)
uploaded_file = st.file_uploader("Upload Excel (.xls)", type=["xls", "xlsx"], label_visibility="collapsed")

# 데이터 초기값
v_3m, v_3y, v_10y = "", "", ""

# 파일 처리 로직 (최소한의 연산)
if uploaded_file:
    try:
        # 헤더 포함 읽기
        df = pd.read_excel(uploaded_file)
        if len(df) > 0:
            v_3m = str(df.iloc[0, 4])   # E2
            v_3y = str(df.iloc[0, 11])  # L2
            v_10y = str(df.iloc[0, 15]) # P2
    except:
        st.error("File Error")

# 5. [핵심] 폼(Form) 사용 - 입력 중 새로고침 방지 (속도 향상)
with st.form("rate_form"):
    st.caption("Enter Rates & Receiver")
    
    # 모바일 세로 스크롤 효율을 위해 컬럼(column) 제거 -> 위에서 아래로 배치
    receiver = st.text_input("To", value=default_receiver)
    
    cd = st.text_input("CD (%)", placeholder="Direct Input")
    k3m = st.text_input("KTB 3M", value=v_3m)
    k3y = st.text_input("KTB 3Y", value=v_3y)
    k10y = st.text_input("KTB 10Y", value=v_10y)

    # 전송 버튼 (폼 안에 있어서 클릭 시에만 전체 코드 실행됨)
    submitted = st.form_submit_button("Send Mail", type="primary")

# 6. 전송 로직
if submitted:
    if not (sender_email and sender_password and receiver):
        st.error("Check Secrets/Receiver")
    elif not (cd and k3m and k3y and k10y):
        st.warning("Enter all rates")
    else:
        try:
            msg = EmailMessage()
            msg['Subject'] = f"[Report] Market Rates {pd.Timestamp.now().strftime('%Y-%m-%d')}"
            msg['From'] = sender_email
            msg['To'] = receiver
            
            body = f"""Market Rates Report:

- CD: {cd}%
- KTB 3M: {k3m}%
- KTB 3Y: {k3y}%
- KTB 10Y: {k10y}%
"""
            msg.set_content(body)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender_email, sender_password)
                smtp.send_message(msg)
            
            # 풍선 효과 제거 -> 심플한 성공 메시지
            st.success("Sent!")
            
        except Exception as e:
            st.error(f"Error: {e}")
