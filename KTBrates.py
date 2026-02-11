import streamlit as st
from email.message import EmailMessage
import gc

# 1. 페이지 설정
st.set_page_config(page_title="Rates", layout="centered")

# 2. 엑셀 데이터 읽기 (기존과 동일, 캐싱 적용)
@st.cache_data(max_entries=1, show_spinner=False)
def get_rates(uploaded_file):
    v_3m, v_3y, v_10y = "", "", ""
    try:
        filename = uploaded_file.name.lower()
        if filename.endswith('.xls'):
            import xlrd
            file_data = uploaded_file.read()
            wb = xlrd.open_workbook(file_contents=file_data)
            sheet = wb.sheet_by_index(0)
            v_3m = str(sheet.cell_value(1, 4))
            v_3y = str(sheet.cell_value(1, 11))
            v_10y = str(sheet.cell_value(1, 15))
            del file_data, wb, sheet
        else:
            import openpyxl
            wb = openpyxl.load_workbook(uploaded_file, data_only=True, read_only=True)
            sheet = wb.active
            v_3m = str(sheet.cell(row=2, column=5).value)
            v_3y = str(sheet.cell(row=2, column=12).value)
            v_10y = str(sheet.cell(row=2, column=16).value)
            wb.close()
            del wb, sheet
    except:
        pass
    gc.collect()
    def clean(v): return v if v and v != "None" else ""
    return clean(v_3m), clean(v_3y), clean(v_10y)

# 3. 설정 로드
secrets = st.secrets.get("gmail", {})
sid = secrets.get("id", "")
spw = secrets.get("pw", "")
srcv = secrets.get("receiver", "")

# 4. UI 구성
uploaded_file = st.file_uploader("Excel", type=["xls", "xlsx"], label_visibility="collapsed")

if uploaded_file:
    v_3m, v_3y, v_10y = get_rates(uploaded_file)
else:
    v_3m, v_3y, v_10y = "", "", ""

with st.form("main"):
    rcv = st.text_input("To", value=srcv, placeholder="Receiver")
    cd = st.text_input("CD (%)", placeholder="Enter CD")
    c1, c2, c3 = st.columns(3)
    k3m = c1.text_input("3M", value=v_3m)
    k3y = c2.text_input("3Y", value=v_3y)
    k10y = c3.text_input("10Y", value=v_10y)
    
    # 엑셀 첨부 여부 선택 (기본값 True)
    attach_excel = st.checkbox("Attach Excel File", value=True)
    
    submitted = st.form_submit_button("Send", type="primary", use_container_width=True)

# 5. 전송 로직 (엑셀 생성 + 첨부)
if submitted:
    if not (sid and spw and rcv):
        st.error("Check Secrets")
    elif not (cd and k3m and k3y and k10y):
        st.warning("Input Data")
    else:
        try:
            # 필요한 라이브러리 지연 로딩 (속도 최적화)
            import smtplib
            import io  # 메모리 파일 처리용
            import openpyxl # 엑셀 생성용
            from datetime import datetime
            
            today_str = datetime.now().strftime('%Y-%m-%d')
            
            # 메일 객체 생성
            msg = EmailMessage()
            msg['Subject'] = f"[Rate] {today_str}"
            msg['From'] = sid
            msg['To'] = rcv
            msg.set_content(f"Market Rates ({today_str})\n\nCD: {cd}%\n3M: {k3m}%\n3Y: {k3y}%\n10Y: {k10y}%")

            # [핵심] 엑셀 파일 생성 및 첨부
            if attach_excel:
                # 1. 메모리 상에 엑셀 워크북 생성
                wb_new = openpyxl.Workbook()
                ws_new = wb_new.active
                ws_new.title = "Market Rates"
                
                # 2. 헤더와 데이터 작성
                headers = ["Date", "CD (%)", "KTB 3M (%)", "KTB 3Y (%)", "KTB 10Y (%)"]
                values = [today_str, cd, k3m, k3y, k10y]
                ws_new.append(headers)
                ws_new.append(values)
                
                # 3. 메모리 버퍼(BytesIO)에 저장 (디스크 저장 X)
                excel_buffer = io.BytesIO()
                wb_new.save(excel_buffer)
                excel_buffer.seek(0) # 포인터를 처음으로 되돌림
                file_data = excel_buffer.read()
                
                # 4. 메일에 파일 첨부
                msg.add_attachment(
                    file_data,
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    filename=f"Rates_{today_str}.xlsx"
                )

            # 전송
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sid, spw)
                smtp.send_message(msg)
            
            st.success("Sent with Excel!" if attach_excel else "Sent!")
            
        except Exception as e:
            st.error(f"Error: {e}")
