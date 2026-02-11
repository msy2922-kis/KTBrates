import streamlit as st
from email.message import EmailMessage
import gc

# 1. 페이지 설정 (제목 없음, 레이아웃 centered)
st.set_page_config(page_title="Rates", layout="centered")

# 2. 엑셀 데이터 추출 함수 (캐시 + 메모리 최적화 + 포맷 자동 인식)
@st.cache_data(max_entries=1, show_spinner=False)
def get_rates(uploaded_file):
    v_3m, v_3y, v_10y = "", "", ""
    
    try:
        filename = uploaded_file.name.lower()
        
        # [Case A] .xls 파일 (구버전) -> xlrd 사용
        if filename.endswith('.xls'):
            import xlrd
            file_data = uploaded_file.read()
            wb = xlrd.open_workbook(file_contents=file_data)
            sheet = wb.sheet_by_index(0)
            
            # 0-based index: Row 2 -> 1 / E->4, L->11, P->15
            v_3m = str(sheet.cell_value(1, 4))
            v_3y = str(sheet.cell_value(1, 11))
            v_10y = str(sheet.cell_value(1, 15))
            
            del file_data, wb, sheet

        # [Case B] .xlsx 파일 (신버전) -> openpyxl 사용
        else:
            import openpyxl
            # read_only=True로 메모리 절약
            wb = openpyxl.load_workbook(uploaded_file, data_only=True, read_only=True)
            sheet = wb.active
            
            # 1-based index: Row 2 -> 2 / E->5, L->12, P->16
            v_3m = str(sheet.cell(row=2, column=5).value)
            v_3y = str(sheet.cell(row=2, column=12).value)
            v_10y = str(sheet.cell(row=2, column=16).value)
            
            wb.close()
            del wb, sheet

    except Exception:
        pass
    
    # 메모리 강제 청소
    gc.collect()
    
    # None 값이 들어오면 빈칸으로 처리
    def clean(v): return v if v and v != "None" else ""
    return clean(v_3m), clean(v_3y), clean(v_10y)

# 3. 설정 로드
secrets = st.secrets.get("gmail", {})
sid = secrets.get("id", "")
spw = secrets.get("pw", "")
srcv = secrets.get("receiver", "")

# 4. UI 구성 (헤드라인 없음)
# 파일 업로더 (라벨 숨김)
uploaded_file = st.file_uploader("Excel", type=["xls", "xlsx"], label_visibility="collapsed")

# 데이터 로딩
if uploaded_file:
    v_3m, v_3y, v_10y = get_rates(uploaded_file)
else:
    v_3m, v_3y, v_10y = "", "", ""

# 5. 입력 폼 (초경량)
with st.form("main"):
    # 수신인
    rcv = st.text_input("To", value=srcv, placeholder="Receiver")
    
    # CD 금리
    cd = st.text_input("CD (%)", placeholder="Enter CD")
    
    # 3M, 3Y, 10Y (가로 3열 배치)
    c1, c2, c3 = st.columns(3)
    k3m = c1.text_input("3M", value=v_3m)
    k3y = c2.text_input("3Y", value=v_3y)
    k10y = c3.text_input("10Y", value=v_10y)

    # 꽉 찬 전송 버튼
    submitted = st.form_submit_button("Send", type="primary", use_container_width=True)

# 6. 전송 로직
if submitted:
    if not (sid and spw and rcv):
        st.error("Check Secrets")
    elif not (cd and k3m and k3y and k10y):
        st.warning("Input Data")
    else:
        try:
            import smtplib
            from datetime import datetime
            
            msg = EmailMessage()
            msg['Subject'] = f"[Rate] {datetime.now().strftime('%Y-%m-%d')}"
            msg['From'] = sid
            msg['To'] = rcv
            msg.set_content(f"CD: {cd}%\n3M: {k3m}%\n3Y: {k3y}%\n10Y: {k10y}%")

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sid, spw)
                smtp.send_message(msg)
            
            st.success("Sent")
            
        except Exception as e:
            st.error("Error")
