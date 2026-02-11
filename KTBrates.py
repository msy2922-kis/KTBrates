import streamlit as st
from email.message import EmailMessage
import gc

# ==========================================
# [사용자 설정] 여기만 수정하세요!
# ==========================================

# 1. 메일 제목 뒤에 항상 붙을 문구
# 예: " (일일보고)" -> "[Rate] 2026-02-11 (일일보고)"
FIXED_TITLE_SUFFIX = "[한국투자증권] 금리데이터송부" 

# 2. 메일 본문 맨 아래에 들어갈 서명/연락처
FIXED_FOOTER = """
감사합니다.

한국투자증권 FICC Sales부

송인호 부서장 02-3276-5318
권서연 대  리 02-3276-5472
문찬호 대  리 02-3276-6496
김종선 대  리 02-3276-5472
진태영 사  원 02-3276-4976

부서 Email: A07910@koreainvestment.com
"""
# ==========================================


# 1. 페이지 설정
st.set_page_config(page_title="Rates", layout="centered")

# 2. 엑셀 데이터 읽기 (최적화된 로직 유지)
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

# 4. UI 구성 (입력창 최소화)
uploaded_file = st.file_uploader("Excel", type=["xls", "xlsx"], label_visibility="collapsed")

if uploaded_file:
    v_3m, v_3y, v_10y = get_rates(uploaded_file)
else:
    v_3m, v_3y, v_10y = "", "", ""

with st.form("main"):
    # 수신인
    rcv = st.text_input("To", value=srcv, placeholder="Receiver")
    
    # 금리 데이터 입력 (제목/내용 입력창 제거됨)
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
            
            # [제목 구성] 날짜 + 고정 문구
            msg['Subject'] = f"{FIXED_TITLE_SUFFIX}{today_str}"
            
            msg['From'] = sid
            msg['To'] = rcv
            
            # [본문 구성] 금리 데이터 + 고정 하단(서명)
            body_txt = f"""Market Rates ({today_str})

CD: {cd}%
3M: {k3m}%
3Y: {k3y}%
10Y: {k10y}%

---
{FIXED_FOOTER}"""
            msg.set_content(body_txt)

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

                # 데이터 생성 (1행부터 시작, 세로형)
                data_rows = [
                    [file_date_str, "CD수익률", val_cd],
                    [file_date_str, "KTB3m", val_3m],
                    [file_date_str, "KTB3y", val_3y],
                    [file_date_str, "KTB10y", val_10y]
                ]
                
                for row in data_rows:
                    ws_new.append(row)

                # 서식 지정 (1~4행의 C열)
                for r in range(1, 5): 
                    cell = ws_new.cell(row=r, column=3)
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
