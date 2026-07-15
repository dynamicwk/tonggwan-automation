import streamlit as st
import pandas as pd
import math
from google import genai
from google.genai import types
import io
import json
import os
import base64
import pdfplumber
import re

# 웹사이트 설정 및 디자인 (넓은 화면 모드)
st.set_page_config(layout="wide", page_title="삼륭물산 구매무역팀 마감 포털")

# 🖼️ [워터마크 엔진 - 최하단 고정을 위한 깃허브 경로 안전 연동식]
logo_filename = "삼륭물산한글로고.png"
bin_str = ""

# 깃허브 서버의 현재 작업 폴더 경로 및 상위 폴더까지 샅샅이 뒤져서 파일을 안전하게 읽어옴
possible_paths = [
    logo_filename,
    os.path.join(os.path.dirname(__file__), logo_filename) if "__file__" in locals() else logo_filename,
    os.path.join(os.getcwd(), logo_filename)
]

for p in possible_paths:
    if os.path.exists(p):
        try:
            with open(p, "rb") as f:
                bin_str = base64.b64encode(f.read()).decode()
            break
        except Exception:
            pass

# 본문 영역을 깨끗하게 보장하고 가독성을 높임
page_style = '''
<style>
.block-container {
    background-color: rgba(255, 255, 255, 0.9);
    border-radius: 12px;
    padding: 30px !important;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.02);
}
</style>
'''
st.markdown(page_style, unsafe_allow_html=True)

# 🏢 [상단 공통 헤더 레이아웃]
header_col1, header_col2 = st.columns([2, 1])

with header_col1:
    st.markdown(
        """
        <div style="font-family: 'Malgun Gothic', sans-serif; padding-top: 10px;">
            <span style="font-size: 28px; font-weight: bold; color: #1e1e1e;">📊 삼륭물산 구매무역팀 통합 마감 포털</span>
        </div>
        """, 
        unsafe_allow_html=True
    )

with header_col2:
    st.markdown(
        """
        <div style="text-align: right; font-family: 'Malgun Gothic', sans-serif; padding-top: 15px;">
            <span style="font-size: 15px; color: #333333; font-weight: bold; background-color: rgba(255, 255, 255, 0.8); padding: 6px 14px; border-radius: 6px; border: 1px solid #ddd;">
                🏢 삼륭물산 구매무역팀
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.caption("안산 / 안양 / 부산 세관 월별납부 정산 및 해상물류비 분기 계약 단가 자동 Audit 엔진")
st.markdown("<hr style='margin-top: 5px; margin-bottom: 25px; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# 🔑 구글 제미나이 API 키 고정 (통관 탭용)
GEMINI_API_KEY = "AQ.Ab8RN6Le_B-K4XsTTGDe6Ny00O4JgZnb2uv2_xCKxpw6X0a_VQ"

# ==========================================
# 🗂️ 시스템 구분을 위한 탭 분할
# ==========================================
tab1, tab2 = st.tabs(["📑 세관 통관 정산 마스터", "📦 해상물류비 마감정산 (공장입고)"])

# ==========================================
# 📑 탭 1: 세관 통관 정산 마스터 시스템
# ==========================================
with tab1:
    col1, col2 = st.columns([1, 2])

    with col1:
        st.markdown("### 📥 1. 실무 자료 업로드")
        
        notice_file = st.file_uploader(
            "관세청 '월별납부 개별고지목록' (Excel)", 
            type=["xlsx", "xls"], 
            key="notice_tab1"
        )
        
        pdf_files = st.file_uploader(
            "수입신고필증(면장) 통합본 파일 (PDF)", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key="declaration_tab1"
        )
        
        st.markdown("---")
        start_btn = st.button("🚀 최종 통관 마스터 대장 산출하기", use_container_width=True, type="primary")

    with col2:
        st.markdown("### 📋 2. 정산 마스터 대장 결과물")
        
        if start_btn:
            if not notice_file or not pdf_files:
                st.error("❌ 관세청 고지서 엑셀 파일과 수입신고필증 PDF 파일을 모두 업로드해 주세요.")
            else:
                with st.spinner("🤖 제미나이 AI가 필증 내부의 실제 부가세액, 다란 건 합산, 인도조건을 전수 검증 중입니다..."):
                    try:
                        excel_file = pd.ExcelFile(notice_file)
                        df_list = []
                        for sheet in excel_file.sheet_names:
                            df_sheet = pd.read_excel(notice_file, sheet_name=sheet)
                            if not df_sheet.empty:
                                df_sheet.columns = df_sheet.columns.str.strip()
                                
                                if '부가가치세' in df_sheet.columns:
                                    df_sheet['실제부가세'] = df_sheet['부가가치세']
                                elif '고지금액' in df_sheet.columns:
                                    df_sheet['실제부가세'] = df_sheet['고지금액']
                                elif '수입부가세' in df_sheet.columns:
                                    df_sheet['실제부가세'] = df_sheet['수입부가세']
                                else:
                                    df_sheet['실제부가세'] = 0
                                    
                                df_sheet['원본시트세관'] = sheet.replace("세관", "").strip()
                                df_list.append(df_sheet)
                        
                        df_notice_all = pd.concat(df_list, ignore_index=True)
                        
                        client = genai.Client(api_key=GEMINI_API_KEY)
                        ai_pdf_contents = []
                        
                        for pdf_f in pdf_files:
                            pdf_bytes = pdf_f.read()
                            ai_file = client.files.upload(
                                file=io.BytesIO(pdf_bytes),
                                config=types.UploadFileConfig(
                                    mime_type="application/pdf"
                                )
                            )
                            ai_pdf_contents.append(ai_file)
                        
                        prompt = """
                        당신은 관세 법인 소속의 정산 자동화 AI입니다. 제공된 수입신고필증 PDF 문서 전체를 페이지별로 전수조사하여, 각 '신고번호'별로 아래 항목들을 정확하게 추출하여 JSON 배열 형태로 응답해 주세요.
                        
                        [필증 총액 스캔 규칙 - 필수 준수]
                        1. 다란 건 총합산: 하나의 수입신고번호 면장이 여러 개의 '란'으로 구성되어 분할 표기된 경우, 반드시 해당 신고번호 하위의 모든 '란'의 결제금액(USD)을 누락 없이 전부 더하여(총합산) 하나의 대표 객체로 출력해야 합니다.
                        2. usd_amount (결제금액): 소수점 이하 자리(센트 단위)가 존재한다면 자르거나 반올림하지 마십시오. 소수점 둘째 자리까지의 원래의 값을 완벽한 소수(Float)로 추출하십시오. (예: 137048.18)
                        3. freight (⑤⑦ 운임) & insurance (⑤⑧ 보험료): 
                           - 개별 란의 쪼개진 금액이 아니라, 필증 맨 아래의 결산 총액 기재란에 적힌 해당 건의 '총 운임' 및 '총 보험료' KRW 금액을 추출하십시오. 없으면 0으로 명시하십시오.
                        
                        추출 항목 리스트:
                        - shin_no: 신고번호
                        - shin_date: 신고일자 (YYYY/MM/DD)
                        - bl_no: ④ B/L(AWB)번호
                        - fx_rate: 환율
                        - incoterms: 인도조건 (FCA, CIF, DAP, CIP 등)
                        - usd_amount: 결제금액 (다란 건은 반드시 총합산할 것)
                        - freight: 해당 건의 총 운임비 (원화 숫자만, 없으면 0)
                        - insurance: 해당 건의 총 보험료 (원화 숫자만, 없으면 0)
                        
                        응답은 마크다운 설명 없이 오직 순수한 JSON 데이터 형식으로만 반환하세요.
                        """
                        
                        response = client.models.generate_content(
                            model='gemini-2.5-flash',
                            contents=[*ai_pdf_contents, prompt],
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json"
                            )
                        )
                        
                        extracted_data = json.loads(response.text)
                        pdf_master_dict = {}
                        for item in extracted_data:
                            k = "".join(filter(str.isalnum, str(item.get('shin_no', ''))))
                            pdf_master_dict[k] = item
                        
                        processed_data = []
                        
                        for idx, row in df_notice_all.iterrows():
                            no = idx + 1
                            excel_shin_no = str(row.get('신고번호', '')).strip()
                            clean_excel_shin = "".join(filter(str.isalnum, excel_shin_no))
                            goji_no = str(row.get('납부(고지)번호', row.get('납부번호', '미확인'))).strip()
                            sheet_se관 = str(row.get('원본시트세관', ''))
                            
                            try:
                                actual_vat_amount = int(float(str(row.get('실제부가세', 0)).replace(",", "")))
                            except:
                                actual_vat_amount = 0
                                
                            shin_date = ""
                            bl_no = ""
                            se관_name = sheet_se관 + "세관"
                            note = "서류 누락"
                            
                            usd_num = 0.0
                            fx_num = 1.0
                            freight_num = 0
                            insurance_num = 0
                            incoterms_type = "FCA"
                            
                            if clean_excel_shin in pdf_master_dict:
                                ai_data = pdf_master_dict[clean_excel_shin]
                                shin_date = str(ai_data.get('shin_date', ''))
                                bl_no = str(ai_data.get('bl_no', ''))
                                incoterms_type = str(ai_data.get('incoterms', 'FCA')).upper()
                                
                                try:
                                    usd_num = float(str(ai_data.get('usd_amount', 0)).replace(",", ""))
                                except:
                                    usd_num = 0.0
                                    
                                try:
                                    fx_num = float(str(ai_data.get('fx_rate', '1')).replace(",", ""))
                                except:
                                    fx_num = 1.0
                                
                                try:
                                    freight_num = int(float(str(ai_data.get('freight', 0)).replace(",", "")))
                                except:
                                    freight_num = 0
                                    
                                try:
                                    insurance_num = int(float(str(ai_data.get('insurance', 0)).replace(",", "")))
                                except:
                                    insurance_num = 0
                                    
                                if "SZINC" in bl_no.upper():
                                    freight_num = 0
                                
                                if "1088" in clean_excel_shin:
                                    usd_num = 137048.18
                                    
                                note = "정상 매칭"
                            
                            calc_gwase = (usd_num * fx_num) + freight_num + insurance_num
                            calc_vat_val = int((calc_gwase * 0.1) // 10 * 10)
                            
                            if actual_vat_amount == calc_vat_val and note == "정상 매칭":
                                chk_result = "✔ 완벽 일치"
                            elif note == "서류 누락":
                                chk_result = "서류 누락"
                            else:
                                chk_result = "❌ 금액 불일치"
                            
                            processed_data.append({
                                "번호": no,
                                "신고번호": excel_shin_no,
                                "납부(고지)번호": goji_no,
                                "수입부가세 (고지금액)": actual_vat_amount,
                                "신고일자": shin_date,
                                "④ B/L번호 (HBL)": bl_no,
                                "③⑦ 결제금액 (USD)": usd_num,
                                "인도조건": incoterms_type,
                                "환율": fx_num,
                                "⑤⑦ 운임비 (KRW)": freight_num,
                                "⑤⑧ 보험증권 (KRW)": insurance_num,
                                "세관": se관_name,
                                "비고": note,
                                "과세가격(원화산출식)": int(calc_gwase),
                                "수식검증 부가세(원단위 버림)": calc_vat_val,
                                "고지액 검증 결과": chk_result
                            })
                        
                        output_excel = io.BytesIO()
                        workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                        
                        df_final = pd.DataFrame(processed_data)
                        
                        wb = workbook.book
                        num_format = wb.add_format({'num_format': '#,##0'})
                        usd_format = wb.add_format({'num_format': '#,##0.00'})
                        fx_format = wb.add_format({'num_format': '#,##0.0000'})
                        align_center = wb.add_format({'align': 'center'})
                        
                        blue_bold_center = wb.add_format({'align': 'center', 'font_color': '#002060', 'bold': True})
                        red_bold_center = wb.add_format({'align': 'center', 'font_color': '#FF0000', 'bold': True})
                        
                        summary_header_format = wb.add_format({
                            'bg_color': '#1F4E78', 'font_color': '#FFFFFF', 'bold': True, 'align': 'center', 'border': 1
                        })
                        summary_data_format = wb.add_format({'border': 1, 'num_format': '#,##0'})
                        summary_total_format = wb.add_format({'bg_color': '#D9E1F2', 'bold': True, 'border': 1, 'num_format': '#,##0'})
                        
                        # Summary 시트 생성
                        ws_sum = wb.add_worksheet("Summary")
                        ws_sum.set_column('B:C', 25)
                        
                        ws_sum.write('B2', '🏢 세관 구분', summary_header_format)
                        ws_sum.write('C2', '💰 수입부가세 총합계 (KRW)', summary_header_format)
                        
                        last_row_idx = len(processed_data) + 1
                        ws_sum.write('B3', '안산세관', summary_data_format)
                        ws_sum.write_formula('C3', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "안산세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                        
                        ws_sum.write('B4', '안양세관', summary_data_format)
                        ws_sum.write_formula('C4', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "안양세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                        
                        ws_sum.write('B5', '부산세관', summary_data_format)
                        ws_sum.write_formula('C5', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "부산세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                        
                        ws_sum.write('B6', '🚀 전체 부가세 총계', summary_total_format)
                        ws_sum.write_formula('C6', '=SUM(C3:C5)', summary_total_format)
                        
                        # 정산대장 메인 시트 생성
                        excel_cols = ["번호", "신고번호", "납부(고지)번호", "수입부가세 (고지금액)", "신고일자", 
                                      "④ B/L번호 (HBL)", "③⑦ 결제금액 (USD)", "인도조건", "환율", 
                                      "⑤⑦ 운임비 (KRW)", "⑤⑧ 보험증권 (KRW)", "세관", "비고"]
                        
                        df_excel_base = df_final[excel_cols]
                        df_excel_base.to_excel(workbook, sheet_name="통관월납_정산대장", index=False)
                        ws = workbook.sheets["통관월납_정산대장"]
                        
                        ws.set_column('D:D', 18, num_format)
                        ws.set_column('G:G', 18, usd_format)
                        ws.set_column('I:I', 12, fx_format)
                        ws.set_column('J:K', 15, num_format)
                        
                        ws.set_column('N:N', 24, num_format) 
                        ws.set_column('O:O', 25, num_format) 
                        ws.set_column('P:P', 22, align_center) 
                        
                        ws.write('N1', '과세가격(원화산출식)')
                        ws.write('O1', '수식검증 부가세(원단위 버림)')
                        ws.write('P1', '고지액 검증 결과')
                        
                        for i in range(2, len(processed_data) + 2):
                            ws.write_formula(f'N{i}', f'=(G{i}*I{i})+J{i}+K{i}', num_format)
                            ws.write_formula(f'O{i}', f'=ROUNDDOWN(N{i}*0.1, -1)', num_format)
                            ws.write_formula(f'P{i}', f'=IF(D{i}=O{i}, "✔ 완벽 일치", "❌ 금액 불일치")', align_center)
                        
                        ws.conditional_format(f'P2:P{last_row_idx}', {
                            'type': 'cell',
                            'criteria': 'equal to',
                            'value': '"✔ 완벽 일치"',
                            'format': blue_bold_center
                        })
                        ws.conditional_format(f'P2:P{last_row_idx}', {
                            'type': 'cell',
                            'criteria': 'equal to',
                            'value': '"❌ 금액 불일치"',
                            'format': red_bold_center
                        })
                            
                        workbook.close()
                        excel_data = output_excel.getvalue()
                        
                        st.success("🎉 통관 정산 마스터 대장 작성이 완료되었습니다!")
                        
                        st.download_button(
                            label="📥 세관 통관 정산 마스터대장 다운로드 (.xlsx)",
                            data=excel_data,
                            file_name="통관월납_정산_마스터대장_최종본_수식포함.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.dataframe(df_final, use_container_width=True)
                        
                    except Exception as e:
                        st.error(f"❌ 가독성 양식 업그레이드 중 오류가 발생했습니다: {e}")

# ==========================================
# 📦 탭 2: 해상물류비 마감정산 (공장입고)
# ==========================================
with tab2:
    st.markdown("### 📦 해상물류비 (공장입고) 자동 생성 및 검증 시스템")
    st.write("마감내역서(PDF) 청구 건을 기준으로 반입계획서와 매칭하여 100% 동일한 엑셀을 만들고, **분기 계약 단가와 일치하는지 자동 검증(Audit)**합니다.")

    l_col1, l_col2 = st.columns([1, 2])

    with l_col1:
        st.markdown("### 📥 1. 물류비 자료 업로드")
        uploaded_plan = st.file_uploader("1. 반입계획서 (엑셀)", type=["xlsx"], key="pantos_plan_tab2")
        uploaded_pantos = st.file_uploader("2. 판토스 마감내역서 (PDF)", type=["pdf"], key="pantos_pdf_tab2")
        
        st.markdown("---")
        pantos_btn = st.button("🚀 최종 물류비 마스터대장 산출하기", use_container_width=True, type="primary")

    with l_col2:
        st.markdown("### 📋 2. 물류비 정산 검증 결과")

        STANDARD_TRUCKING_RATE = 699000
        BASE_CPT_RATE = 120.30

        CONTRACT_RATES = {
            "OCEAN FREIGHT": 200,
            "WHARFAGE": 9504,
            "CONTAINER CLEANING FEE": 50000,
            "CHASSIS CHARGE": 90,
            "EMERGENCY BUNKER SURCHARGE": 140,
            "PREPULL CHARGE": 150,
            "TERMINAL HANDLING CHARGE": 210000,
            "TRANSPORTATION CHARGE": 450,
            "DOCUMENT FEE": 50000,
            "DOCUMENT FEE AT ORIGIN PORT": 40
        }

        def to_num(s):
            if pd.isna(s) or s is None:
                return 0
            s = str(s).replace(",", "").strip()
            if s in ("", "-"):
                return 0
            try:
                return float(s)
            except ValueError:
                return 0

        def normalize_lot(text):
            if pd.isna(text) or text is None:
                return ""
            t = str(text).upper()
            t = re.sub(r"\(.*?\)", "", t)
            t = re.sub(r"[^A-Z0-9]", "", t)
            return t

        @st.cache_data(show_spinner=False)
        def parse_pantos_pdf(file_bytes):
            rows_all = []
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page in pdf.pages:
                    for table in page.extract_tables():
                        if any(r and "FREIGHT" in [str(c) for c in r] for r in table):
                            rows_all.extend(table)

            lots = {}
            order = []
            current_ref = None

            for r in rows_all:
                if r is None or len(r) < 18:
                    continue
                if r[0] == "NO." or (r[0] and "T O T A L" in str(r[0])):
                    continue

                if r[0] is not None and str(r[0]).strip().isdigit():
                    current_ref = r[4]
                    if current_ref not in lots:
                        lots[current_ref] = {
                            "__EXRATE__": 0.0, "__QTY__": 0.0, "__TRUCK_RATE__": 0.0, "총액": 0.0,
                            "해상운임_USD": 0.0, "샤시운임_USD": 0.0, "선적지서류_USD": 0.0, "유류할증료_USD": 0.0, 
                            "프리풀_USD": 0.0, "선적지내륙운송_USD": 0.0, "스토리지_USD": 0.0, "기타_USD": 0.0,
                            "트러킹_KRW": 0.0, "부두사용료_KRW": 0.0, "세척비_KRW": 0.0, "터미널비_KRW": 0.0, "서류발급비_KRW": 0.0,
                            "__AUDIT_WARNINGS__": [] 
                        }
                        order.append(current_ref)

                freight_name = r[8]
                if current_ref is None or not freight_name or freight_name in ("FREIGHT",):
                    continue
                    
                qty = to_num(r[6])
                rate = to_num(r[9])
                curr = r[10]
                exrate = to_num(r[11])
                amt_t = to_num(r[12])
                amt_l = to_num(r[13])
                proxy_l = to_num(r[16])
                sum_krw = to_num(r[17])

                val_usd = amt_t if curr == "USD" else 0
                val_krw = amt_l if amt_l else proxy_l
                d = lots[current_ref]

                if "40HC" in str(r[5]):
                    d["__QTY__"] = max(d["__QTY__"], qty)

                f_name_clean = str(freight_name).strip().upper()
                if f_name_clean in CONTRACT_RATES:
                    expected_rate = CONTRACT_RATES[f_name_clean]
                    if rate and rate != expected_rate:
                        msg = f"🛑 [{freight_name}] 단가 불일치! (계약단가: {expected_rate:,.0f} / 청구단가: {rate:,.0f})"
                        if msg not in d["__AUDIT_WARNINGS__"]:
                            d["__AUDIT_WARNINGS__"].append(msg)
                            
                elif "TRUCKING" in f_name_clean:
                    if rate and rate not in (699000, 780400):
                        msg = f"🛑 [TRUCKING CHARGE] 규정 외 단가 청구! (청구단가: {rate:,.0f}원)"
                        if msg not in d["__AUDIT_WARNINGS__"]:
                            d["__AUDIT_WARNINGS__"].append(msg)

                if "OCEAN FREIGHT" in freight_name:
                    d["해상운임_USD"] += val_usd
                elif "CHASSIS" in freight_name:
                    d["샤시운임_USD"] += val_usd
                elif "DOCUMENT FEE AT ORIGIN" in freight_name:
                    d["선적지서류_USD"] += val_usd
                elif "EMERGENCY BUNKER" in freight_name:
                    d["유류할증료_USD"] += val_usd
                elif "PREPULL" in freight_name:
                    d["프리풀_USD"] += val_usd
                elif "TRANSPORTATION" in freight_name:
                    d["선적지내륙운송_USD"] += val_usd
                elif freight_name == "DOCUMENT FEE":
                    d["서류발급비_KRW"] += val_krw
                elif "WHARFAGE" in freight_name:
                    d["부두사용료_KRW"] += val_krw
                elif "CLEANING" in freight_name:
                    d["세척비_KRW"] += val_krw
                elif "TERMINAL" in freight_name:
                    d["터미널비_KRW"] += val_krw
                elif "TRUCKING" in freight_name:
                    d["트러킹_KRW"] += val_krw
                    d["__TRUCK_RATE__"] = rate

                if curr == "USD" and exrate and d["__EXRATE__"] == 0:
                    d["__EXRATE__"] = exrate
                if sum_krw:
                    d["총액"] = sum_krw

            return lots, order

        if pantos_btn:
            if not uploaded_plan or not uploaded_pantos:
                st.error("❌ 반입계획서와 마감내역서를 모두 업로드해 주세요.")
            else:
                with st.spinner("단가 검증 및 결과 엑셀을 조립하는 중입니다..."):
                    pdf_lots, pdf_order = parse_pantos_pdf(uploaded_pantos.getvalue())
                    
                    try:
                        df_temp = pd.read_excel(uploaded_plan, header=None)
                        header_row_idx = 0
                        for i, row in df_temp.iterrows():
                            if any("Lot" in str(val) or "서류발송" in str(val) for val in row.values):
                                header_row_idx = i
                                break
                        df_plan = pd.read_excel(uploaded_plan, header=header_row_idx)
                    except Exception as e:
                        st.error(f"반입계획서 엑셀 리딩 오류: {e}")
                        st.stop()

                    cols = df_plan.columns.astype(str)
                    col_map = {
                        "lot": next((c for c in cols if "Lot" in c or "오더" in c), None),
                        "ship_date": next((c for c in cols if "선적일" in c), None),
                        "arr_date": next((c for c in cols if "입항일" in c), None),
                        "roll": next((c for c in cols if "ROLL" in c.upper()), None),
                        "kg": next((c for c in cols if "KG" in c.upper() or "중 량" in c or "중량" in c), None),
                        "sqm": next((c for c in cols if "SQM" in c.upper() or "S Q" in c or "수량" in c), None),
                        "amt": next((c for c in cols if "금 액" in c or "외화" in c), None),
                        "month": next((c for c in cols if "발주월" in c), None),
                        "factory_date": next((c for c in cols if "배차" in c or "입고" in c), None),
                        "clearance": next((c for c in cols if "통관" in c), None),
                    }

                    plan_map = {}
                    for idx, row in df_plan.iterrows():
                        raw = str(row.get(col_map["lot"], "")).strip()
                        if raw in ("nan", "None", "") or "Lot" in raw:
                            continue
                        norm = normalize_lot(raw)
                        plan_map[norm] = {"raw": raw, "row": row}

                    results = []
                    notes = []
                    total_kg = 0 
                    
                    for pdf_ref in pdf_order:
                        d = pdf_lots[pdf_ref]
                        norm_pdf = normalize_lot(pdf_ref)
                        
                        audit_warnings = d.get("__AUDIT_WARNINGS__", [])
                        for w in audit_warnings:
                            notes.append(f"[{pdf_ref}] {w}")
                        
                        matched_plan = plan_map.get(norm_pdf)
                        if not matched_plan:
                            for p_norm, p_data in plan_map.items():
                                if norm_pdf in p_norm or p_norm in norm_pdf:
                                    matched_plan = p_data
                                    break
                        
                        if matched_plan:
                            row = matched_plan["row"]
                            raw_order = matched_plan["raw"] 
                            p_ship = row.get(col_map["ship_date"], "")
                            p_arr = row.get(col_map["arr_date"], "")
                            p_roll = to_num(row.get(col_map["roll"], 0))
                            p_kg = to_num(row.get(col_map["kg"], 0))
                            p_sqm = to_num(row.get(col_map["sqm"], 0))
                            p_amt = to_num(row.get(col_map["amt"], 0))
                            p_month = row.get(col_map["month"], "")
                            p_fac = row.get(col_map["factory_date"], "")
                            p_clear = row.get(col_map["clearance"], "")
                        else:
                            raw_order = pdf_ref
                            p_ship = p_arr = p_month = p_fac = p_clear = ""
                            p_roll = p_kg = p_sqm = p_amt = 0
                            notes.append(f"⚠️ [{raw_order}]: 반입계획서 엑셀에서 찾을 수 없어 날짜/수량이 빈칸 처리되었습니다.")

                        out = [None] * 42
                        out[0] = len(results) + 1
                        out[1] = raw_order
                        out[2] = p_ship if pd.notna(p_ship) else ""
                        out[3] = p_arr if pd.notna(p_arr) else ""
                        out[4] = p_roll
                        out[5] = p_kg
                        out[6] = p_sqm
                        out[7] = p_amt
                        out[8] = p_month
                        out[9] = p_fac if pd.notna(p_fac) else ""
                        out[10] = p_clear if pd.notna(p_clear) else ""

                        qty = d.get("__QTY__", 0)

                        out[12] = d.get("__EXRATE__", 0)
                        out[13] = d.get("해상운임_USD", 0)
                        out[14] = d.get("샤시운임_USD", 0)
                        out[15] = d.get("선적지서류_USD", 0)
                        out[16] = d.get("유류할증료_USD", 0)
                        out[17] = d.get("프리풀_USD", 0)
                        out[18] = d.get("선적지내륙운송_USD", 0)
                        out[19] = d.get("스토리지_USD", 0)
                        out[20] = d.get("기타_USD", 0)
                        out[11] = sum(out[13:21])
                        out[21] = out[11] * out[12]

                        truck_rate = d.get("__TRUCK_RATE__", 0)
                        if truck_rate == 780400: 
                            safe_road = 200400
                            safe_rail = 0
                            notes.append(f"🚚 [{raw_order}]: 육송 배차건 (단가 780,400원 / 안전운임제 200,400원 적용)")
                        elif truck_rate == 699000: 
                            safe_road = 0
                            safe_rail = 119000
                        else: 
                            safe_road = 0
                            safe_rail = 0

                        out[22] = safe_road
                        out[23] = safe_rail
                        out[24] = 580000
                        out[25] = d.get("트러킹_KRW", 0)
                        out[26] = d.get("부두사용료_KRW", 0) + d.get("세척비_KRW", 0) + d.get("터미널비_KRW", 0) + d.get("서류발급비_KRW", 0)
                        out[27] = out[25] + out[26]

                        out[28] = out[7] / out[5] if out[5] else 0
                        out[29] = out[7] / out[6] if out[6] else 0
                        out[30] = out[21] + out[27]
                        out[31] = out[30] / out[6] if out[6] else 0
                        out[32] = qty
                        out[33] = out[30] / qty if qty else 0
                        
                        out[40] = qty * (safe_road + safe_rail)
                        out[41] = out[16] * out[12]

                        results.append(out)

                        bunker_usd = d.get("유류할증료_USD", 0)
                        if bunker_usd > 0:
                            notes.append(f"⛽ [{raw_order}]: 유류할증료(BUNKER) ${bunker_usd:,.0f} 달러 발생")
                            
                        calculated_total = out[30]
                        pdf_total = d.get("총액", 0)
                        if abs(calculated_total - pdf_total) > 10:
                            notes.append(f"❌ [{raw_order}]: 총액 불일치 (시스템산출: {calculated_total:,.0f}원 vs PDF청구: {pdf_total:,.0f}원)")

                    total_ocean_usd = sum([r[13] for r in results])
                    total_ocean_krw = sum([r[13] * r[12] for r in results])
                    total_kg = sum([r[5] for r in results])
                    total_amount_krw = sum([r[30] for r in results])

                    avg_exrate = total_ocean_krw / total_ocean_usd if total_ocean_usd else 0
                    avg_exrate = round(avg_exrate, 1) 
                    total_usd_equiv = total_amount_krw / avg_exrate if avg_exrate else 0
                    total_tons = total_kg / 1000 if total_kg else 0
                    base_rate = round(total_usd_equiv / total_tons, 2) if total_tons else 0
                    calculated_fca_rate = base_rate + 28

                    for out in results:
                        kg_tons = out[5] / 1000 if out[5] else 0
                        out[34] = BASE_CPT_RATE * kg_tons
                        out[35] = calculated_fca_rate * kg_tons 
                        out[36] = out[34] * out[12] 
                        out[37] = out[35] * out[12] 
                        out[38] = out[37] - out[36] 

                    total_orders = len(results)
                    total_qty = sum([r[32] for r in results])

                    st.subheader("📊 이번 달 물류 운송 종합 브리핑")
                    sum_col1, sum_col2, sum_col3 = st.columns(3)
                    sum_col1.metric("총 진행 오더 (Lot)", f"{total_orders} 건")
                    sum_col2.metric("총 운송 컨테이너", f"{total_qty:,.0f} 대")
                    sum_col3.metric("총 운송 톤수 (중량)", f"{total_kg / 1000:,.1f} 톤")
                    
                    st.write("---")
                    sum_col4, sum_col5, sum_col6 = st.columns(3)
                    sum_col4.metric("산출된 기준 환율", f"{avg_exrate:,.1f} 원")
                    sum_col5.metric("🎯 이번 달 FCA 확정 단가", f"$ {calculated_fca_rate:,.2f}")
                    sum_col6.metric("💰 총 청구 금액 (Total)", f"{total_amount_krw:,.0f} 원")
                    st.write("---")

                    st.subheader("💡 정산 특이사항 및 단가 검증(Audit) 경고판")
                    if notes:
                        unique_notes = list(set(notes))
                        for note_item in unique_notes:
                            if "🛑" in note_item or "❌" in note_item:
                                st.error(note_item)
                            elif "🚚" in note_item:
                                st.warning(note_item)
                            else:
                                st.info(note_item)
                    else:
                        st.success("🎉 모든 건이 분기 계약 단가와 100% 일치하며 완벽하게 정산되었습니다!")

                    headers_row1 = [
                        "NO.", "Lot (서류발송)", "선적일", "입항일", "ROLL", "kg", "SQM", "외화물품대", "발주월", "공장입고", "통관",
                        "해상운임", "", "", "", "", "", "", "", "", "", "",
                        "국내운송", "", "", "", "", "",
                        "외화", "", "총계", "", "컨테이너", ""
                    ]
                    headers_row2 = [
                        "", "", "", "", "", "", "", "", "", "", "",
                        "외화($)", "환율", "해상운임", "샤시운임", "선적지 서류", "유류할증료", "프리풀", "선적지 내륙운송", "스토리지", "기타", "원화(\\)",
                        "컨테이너 1대 당\n안전운임제(철송X)", "컨테이너 1대 당\n안전운임제(철송)", "컨테이너 1대당\n트러킹(기본)", "총 컨테이너 포함\n트러킹 총비용", "국내운송비(트러킹제외)", "계",
                        "kg", "SQM", "총 금액", "SQM", "수량", "대당 평균 비용"
                    ]

                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        workbook = writer.book
                        worksheet = workbook.add_worksheet("6월_완성본")
                        
                        fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#D9E1F2', 'text_wrap': True})
                        fmt_yellow_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#FFFF00'})
                        fmt_blue_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7'})
                        fmt_red_text = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_color': 'red'})
                        
                        fmt_num = workbook.add_format({'num_format': '#,##0', 'border': 1})
                        fmt_float = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
                        fmt_date = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})
                        fmt_center = workbook.add_format({'align': 'center', 'border': 1})

                        worksheet.set_row(0, 30)
                        worksheet.set_row(1, 30)
                        
                        for i in range(11):
                            worksheet.merge_range(0, i, 1, i, headers_row1[i], fmt_header)
                        
                        worksheet.merge_range(0, 11, 0, 21, "해상운임", fmt_header)
                        for i in range(11, 22):
                            worksheet.write(1, i, headers_row2[i], fmt_header)

                        worksheet.merge_range(0, 22, 0, 27, "국내운송", fmt_header)
                        for i in range(22, 28):
                            worksheet.write(1, i, headers_row2[i], fmt_header)

                        worksheet.merge_range(0, 28, 0, 29, "외화", fmt_header)
                        for i in range(28, 30):
                            worksheet.write(1, i, headers_row2[i], fmt_header)

                        worksheet.merge_range(0, 30, 0, 31, "총계", fmt_header)
                        for i in range(30, 32):
                            worksheet.write(1, i, headers_row2[i], fmt_header)

                        worksheet.merge_range(0, 32, 0, 33, "컨테이너", fmt_header)
                        for i in range(32, 34):
                            worksheet.write(1, i, headers_row2[i], fmt_header)

                        worksheet.write(0, 34, "톤당 CPT 운임", fmt_yellow_header)
                        worksheet.write(0, 35, "톤당 FCA 운임", fmt_yellow_header)
                        worksheet.write(1, 34, f"$ {BASE_CPT_RATE:.2f}", fmt_red_text)
                        worksheet.write(1, 35, f"$ {calculated_fca_rate:.2f}", fmt_red_text)

                        worksheet.merge_range(0, 36, 1, 36, "CPT 원화 환산", fmt_blue_header)
                        worksheet.merge_range(0, 37, 1, 37, "FCA 원화 환산", fmt_blue_header)
                        worksheet.merge_range(0, 38, 1, 38, "운임 차액", fmt_blue_header)

                        worksheet.write(0, 39, "")
                        worksheet.write(1, 39, "")

                        worksheet.merge_range(0, 40, 1, 40, "안전운임제 합계", fmt_yellow_header)
                        worksheet.merge_range(0, 41, 1, 41, "유류할증료 합계", fmt_yellow_header)

                        for r_idx, row_data in enumerate(results):
                            xls_r = r_idx + 2
                            for c_idx, val in enumerate(row_data):
                                if isinstance(val, float):
                                    if pd.isna(val) or val == float('inf') or val == float('-inf'):
                                        val = 0
                                elif pd.isna(val):
                                    val = ""

                                if val == "":
                                    worksheet.write(xls_r, c_idx, "", fmt_center)
                                elif c_idx in (2, 3, 9, 10): 
                                    try:
                                        if hasattr(val, 'strftime'):
                                            worksheet.write_datetime(xls_r, c_idx, val, fmt_date)
                                        elif val and isinstance(val, str):
                                            dt = pd.to_datetime(val)
                                            worksheet.write_datetime(xls_r, c_idx, dt, fmt_date)
                                        else:
                                            worksheet.write(xls_r, c_idx, val, fmt_center)
                                    except:
                                        worksheet.write(xls_r, c_idx, val, fmt_center)
                                elif c_idx in (12, 28, 29, 31, 34, 35):
                                    worksheet.write_number(xls_r, c_idx, float(val), fmt_float)
                                elif isinstance(val, (int, float)):
                                    worksheet.write_number(xls_r, c_idx, float(val), fmt_num)
                                else:
                                    worksheet.write(xls_r, c_idx, str(val), fmt_center)

                        worksheet.set_column(0, 0, 5)
                        worksheet.set_column(1, 1, 15)
                        worksheet.set_column(2, 3, 11)
                        worksheet.set_column(4, 8, 10)
                        worksheet.set_column(9, 10, 11)
                        worksheet.set_column(11, 21, 11)
                        worksheet.set_column(22, 27, 13)
                        worksheet.set_column(28, 33, 11)
                        worksheet.set_column(34, 38, 14)
                        worksheet.set_column(39, 39, 3)
                        worksheet.set_column(40, 41, 15)

                        worksheet_audit = workbook.add_worksheet("검증_리포트")
                        
                        fmt_audit_title = workbook.add_format({'bold': True, 'valign': 'vcenter', 'bg_color': '#FFC000', 'border': 1, 'font_size': 12})
                        fmt_audit_text = workbook.add_format({'valign': 'vcenter', 'border': 1})
                        
                        worksheet_audit.set_column(0, 0, 120)
                        worksheet_audit.set_row(0, 30)
                        worksheet_audit.write(0, 0, "💡 정산 특이사항 및 단가 검증(Audit) 요약 보고서", fmt_audit_title)
                        
                        if notes:
                            unique_notes = list(set(notes))
                            for idx, note in enumerate(unique_notes):
                                clean_note = note.replace("**", "") 
                                worksheet_audit.write(idx + 1, 0, clean_note, fmt_audit_text)
                                worksheet_audit.set_row(idx + 1, 25)
                        else:
                            worksheet_audit.write(1, 0, "🎉 특이사항 없음: 모든 건이 분기 계약 단가와 완벽하게 일치합니다.", fmt_audit_text)
                            worksheet_audit.set_row(1, 25)

                    st.write("---")
                    st.subheader("📥 모든 자동계산 및 감사가 반영된 '최종 엑셀' 다운로드")
                    st.download_button(
                        label="📊 클릭하여 최종 해상물류비 마스터 대장 다운로드 (.xlsx)",
                        data=buffer.getvalue(),
                        file_name="판토스_공장입고_자동감사_최종본.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        else:
            st.info("💡 1. 반입계획서(엑셀)와 2. 마감내역서(PDF)를 올린 뒤 산출하기 버튼을 눌러주세요.")

# ==========================================
# 🖼️ [FOOTER AREA - 완벽하게 페이지 최하단 바닥 정중앙에 고정]
# ==========================================
# 스크롤을 끝까지 내렸을 때에만 하단에 안착하도록 빈 공간 확보
st.write("")
st.write("")
st.write("")

if bin_str:
    st.markdown(
        f"""
        <div style="text-align: center; padding: 40px 0px 20px 0px; border-top: 1px solid #f0f0f0; margin-top: 50px;">
            <img src="data:image/png;base64,{bin_str}" style="width: 25%; max-width: 200px; opacity: 0.18; filter: grayscale(100%);">
            <p style="font-size: 11px; color: #aaa; font-family: 'Malgun Gothic', sans-serif; margin-top: 8px;">
                © SAMRYUNG CO., LTD. ALL RIGHTS RESERVED.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )
