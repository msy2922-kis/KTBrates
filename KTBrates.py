import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage

# 페이지 설정
st.set_page_config(page_title="금리 데이터 전송 (Gmail)", layout="centered")

# --- [비밀번호 관리] st.secrets에서 정보 불러오기 ---
secrets = st.secrets.get("gmail", {})
default_sender = secrets.get("id", "")
default_password = secrets.get("pw", "")
default_receiver = secrets.get("receiver", "")

st.title("📧 Gmail 금리 데이터 전송")
st.markdown("---")

# 1. 계정 설정 (비밀번호가 있으면 접어둠)
with st.expander("🔑 Gmail 계정 설정 (자동 입력됨)", expanded=(not default_sender)):
    col_id, col_pw = st.columns(2)
    with col_id:
        sender_email = st.text_input("Gmail 주소", value=default_sender, placeholder="example@gmail.com")
    with col_pw:
        sender_password = st.text_input("Gmail 앱 비밀번호", value=default_password, type="password", help="구글 계정에서 발송받은 16자리 앱 비밀번호")

# 2. 수신인 설정
receiver_email = st.text_input("📩 받는 사람 이메일 주소", value=default_receiver, placeholder="receiver@example.com")

st.markdown("---")

# 3. 엑셀 파일 업로드 (자동 입력 기능)
st.subheader("📂 데이터 업로드")
uploaded_file = st.file_uploader("채권시가평가기준수익률 엑셀 파일을 업로드하세요", type=["xlsx", "xls"])

# 변수 초기화 (기본값은 빈칸)
val_cd, val_3m, val_3y, val_10y = "", "", "", ""

# 파일이 업로드되면 특정 셀(E2, L2, P2)에서 값 추출
if uploaded_file:
    try:
        # 엑셀 읽기 (첫 번째 시트, 헤더는 첫 줄로 가정)
        df = pd.read_excel(uploaded_file)
        
        # 데이터가 있는 첫 번째 행(Excel의 2행)을 가져옴 -> index 0
        # E열(5번째) -> index 4
        # L열(12번째) -> index 11
        # P열(16번째) -> index 15
        
        # 값이 없을 수도 있으므로 안전하게 가져오기
        if len(df) > 0:
            val_3m = str(df.iloc[0, 4])   # E2
            val_3y = str(df.iloc[0, 11])  # L2
            val_10y = str(df.iloc[0, 15]) # P2
            st.success("✅ 엑셀 파일(E2, L2, P2)에서 금리 정보를 가져왔습니다!")
        else:
            st.warning("엑셀 파일에 데이터가 없습니다.")
            
    except Exception as e:
        st.error(f"엑셀 읽기 실패: {e}")

# 4. 금리 데이터 확인 및 수정
st.subheader("📈 금리 정보 확인")
st.info("CD 수익률은 직접 입력해주세요. 나머지는 자동 입력됩니다.")

c1, c2 = st.columns(2)
with c1:
    # CD는 자동 입력 대상이 아니므로 빈칸(또는 이전 입력값) 유지
    final_cd = st.text_input("CD수익률 (%)", value=val_cd, placeholder="직접 입력 (예: 3.50)")
    final_3m = st.text_input("KTB 3M (%)", value=val_3m, placeholder="E2 셀 값")
with c2:
    final_3y = st.text_input("KTB 3Y (%)", value=val_3y, placeholder="L2 셀 값")
    final_10y = st.text_input("KTB 10Y (%)", value=val_10y, placeholder="P2 셀 값")

# 5. 전송 버튼
if st.button("🚀 Gmail로 전송", use_container_width=True):
    if not (sender_email and sender_password and receiver_email):
        st.warning("이메일 계정 정보를 모두 입력해주세요.")
    elif not (final_cd and final_3m and final_3y and final_10y):
        st.warning("모든 금리 데이터(CD 포함)를 입력해주세요.")
    else:
        try:
            # 메일 객체 생성
            msg = EmailMessage()
            msg['Subject'] = f"📊 [시장금리 보고] {pd.Timestamp.now().strftime('%Y-%m-%d')}"
            msg['From'] = sender_email
            msg['To'] = receiver_email
            
            body = f"""안녕하세요, 금일 시장금리 현황을 보고드립니다.

- CD수익률: {final_cd}%
- KTB 3M: {final_3m}%
- KTB 3Y: {final_3y}%
- KTB 10Y: {final_10y}%

본 메일은 시스템에 의해 자동 발송되었습니다.
"""
            msg.set_content(body)

            # Gmail SMTP 서버 설정
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sender_email, sender_password)
                smtp.send_message(msg)
            
            st.balloons()
            st.success(f"✅ {receiver_email} 님에게 메일을 성공적으로 보냈습니다!")
        except Exception as e:
            st.error(f"❌ 발송 실패: {e}")
