import streamlit as st
from email.message import EmailMessage
import gc

# 1. 페이지 설정
st.set_page_config(page_title="Rates", layout="centered")

# 2. 엑셀 데이터 읽기
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
    
    attach_excel = st.checkbox("Attach Excel File", value=True)
    submitted = st.form_submit_button("Send", type="primary", use_container_width=True)

# 5. 전송 로직
if submitted:
    if not (sid and spw and rcv):
        st.error("Check Secrets")
    elif not (cd and k3m and k3y and k10y):
        st.warning("Input Data")
    else:
        try:
            import smtplib
            import io
            import openpyxl
            from datetime import datetime
            
            now = datetime.now()
            today_str = now.strftime('%Y-%m-%d')
            file_date_str = now.strftime('%Y%m%d')
            
            msg = EmailMessage()
            msg['Subject'] = f"[Rate] {today_str}"
            msg['From'] = sid
            msg['To'] = rcv
            msg.set_content(f"Market Rates ({today_str})\n\nCD: {cd}%\n3M: {k3m}%\n3Y: {k3y}%\n10Y: {k10y}%")

            if attach_excel:
                wb_new = openpyxl.Workbook()
                ws_new = wb_new.active
                ws_new.title = "Sheet1"
                
                try:
                    val_cd = float(cd)
                    val_3m = float(k3m)
                    val_3y = float(k3y)
                    val_10y = float(k10y)
                except ValueError:
                    val_cd, val_3m, val_3y, val_10y = cd, k3m, k3y, k10y

                # --- [헤더 삭제됨] 바로 데이터 시작 ---
                data_rows = [
                    [file_date_str, "CD수익률", val_cd],
                    [file_date_str, "KTB3m", val_3m],
                    [file_date_str, "KTB3y", val_3y],
                    [file_date_str, "KTB10y", val_10y]
                ]
                
                for row in data_rows:
                    ws_new.append(row)

                # 서식 지정: 이제 1행부터 4행까지 데이터가 존재함
                # range(1, 5) -> 1, 2, 3, 4행
                for r in range(1, 5): 
                    cell = ws_new.cell(row=r, column=3) # C열
                    cell.number_format = '0.00'

                excel_buffer = io.BytesIO()
                wb_new.save(excel_buffer)
                excel_buffer.seek(0)
                file_data = excel_buffer.read()
                
                file_name = f"금리data_{file_date_str}.xlsx"
                
                msg.add_attachment(
                    file_data,
                    maintype='application',
                    subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    filename=file_name
                )

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(sid, spw)
                smtp.send_message(msg)
            
            st.success(f"Sent! ({file_name})" if attach_excel else "Sent!")
            
        except Exception as e:
            st.error(f"Error: {e}")
