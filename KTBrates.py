import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage

# 페이지 설정
st.set_page_config(page_title="금리 데이터 전송 (Gmail)", layout="centered")

# --- [비밀번호 관리] st.secrets에서 정보 불러오기 ---
# 웹(Streamlit Cloud)에 배포했을 때나 로컬의 secrets.toml 파일에서 정보를 가져옵니다.
# 정보가 없으면 빈칸("")으로 둡니다.
secrets = st.secrets.get("gmail", {})
default_sender = secrets.get("id", "")
default_password = secrets.get("pw", "")
default_receiver = secrets.get("receiver", "")

st.title("📧 Gmail 금리 데이터 전송")
st.markdown("---")

# 1. Gmail 계정 설정 (평소에는 접어둠)
with st.expander("🔑 Gmail 계정 설정 (자동 입력됨)", expanded=(not default_sender)):
    col_id, col_pw = st.columns(2)
    with col_id:
        sender_email = st.text_input("Gmail 주소", value=default_sender, placeholder="example@gmail.com")
    with col_pw:
        sender_password = st.text_input("Gmail 앱 비밀번호", value=default_password, type="password", help="구글 계정에서 발송받은 16자리 앱 비밀번호")

# 2. 수신인 설정
receiver_email = st.text_input("📩 받는 사람 이메일 주소", value=default_receiver, placeholder="receiver@example.com")

st.markdown("---")

# 3. 금리 데이터 직접 입력 (엑셀 기능 삭제됨)
st.subheader("📈 금리 정보 입력")
st.info("금일 금리 데이터를 직접 입력해주세요.")

c1, c2 = st.columns(2)
with c1:
    val_cd = st.text_input("CD수익률 (%)", placeholder="예: 3.50")
    val_3m = st.text_input("KTB 3M (%)", placeholder="예: 3.45")
with c2:
    val_3y = st.text_input("KTB 3Y (%)", placeholder="예: 3.20")
    val_10y = st.text_input("KTB 10Y (%)", placeholder="예: 3.25")

# 4. 전송 버튼
if st.button("🚀 Gmail로 전송", use_container_width=True):
    # 입력값 검증
    if not (sender_email and sender_password and receiver_email):
        st.warning("이메일 계정 정보(발신인, 비밀번호, 수신인)를 모두 입력해주세요.")
    elif not (val_cd and val_3m and val_3y and val_10y):
        st.warning("금리 데이터를 모두 입력해주세요.")
    else:
        try:
            # 메일 객체 생성
            msg = EmailMessage()
            msg['Subject'] = f"📊 [시장금리 보고] {pd.Timestamp.now().strftime('%Y-%m-%d')}"
            msg['From'] = sender_email
            msg['To'] = receiver_email
            
            body = f"""안녕하세요, 금일 시장금리 현황을 보고드립니다.

- CD수익률: {val_cd}%
- KTB 3M: {val_3m}%
- KTB 3Y: {val_3y}%
- KTB 10Y: {val_10y}%

본 메일은 시스템에 의해 자동 발송되었습니다.
"""
            msg.set_content(body)

            # Gmail SMTP 서버 설정 (SSL 방식)
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender_email, sender_password)
                smtp.send_message(msg)
            
            st.balloons()
            st.success(f"✅ {receiver_email} 님에게 메일을 성공적으로 보냈습니다!")
        except Exception as e:
            st.error(f"❌ 발송 실패: {e}\n\n구글 계정의 '앱 비밀번호(16자리)'를 정확히 입력했는지 확인해주세요.")
