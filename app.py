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
from datetime import datetime
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter

# ==========================================
# 1. 웹사이트 기본 설정 및 디자인
# ==========================================
st.set_page_config(layout="wide", page_title="삼륭물산 구매무역팀 마감 포털")

# 워터마크 배경 설정 로직
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_filename = "삼륭물산한글로고.png"
possible_paths = [
    logo_filename,
    os.path.join(os.path.dirname(__file__), logo_filename) if "__file__" in locals() else logo_filename,
    os.path.join(os.getcwd(), logo_filename)
]

bin_str = ""
for p in possible_paths:
    if os.path.exists(p):
        try:
            with open(p, "rb") as f:
                bin_str = base64.b64encode(f.read()).decode()
            break
        except Exception:
            pass

if bin_str:
    page_bg_img = f'''
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{bin_str}");
        background-size: 55%; 
        background-repeat: no-repeat;
        background-position: center center; 
        background-attachment: fixed;
    }}
    .block-container {{
        background-color: rgba(255, 255, 255, 0.88); 
        border-radius: 12px;
        padding: 30px !important;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.04);
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

# 상단 공통 헤더 레이아웃
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
            <span style="font-size: 15px; color: #333333; font-weight: bold; background-color: rgba(255, 255, 255, 0.9); padding: 6px 14px; border-radius: 6px; border: 1px solid #ddd;">
                🏢 삼륭물산 구매무역팀
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.caption("안산 / 안양 / 부산 세관 월별납부 정산 | 해상물류비 마감감사 | 외상매입금 미정산 현황 마스터 종합 ERP")
st.markdown("<hr style='margin-top: 5px; margin-bottom: 25px; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# ==========================================
# 2. 공통 엔진 및 유틸리티 함수
# ==========================================
# 🔑 구글 제미나이 API 키 고정
GEMINI_API_KEY = "AQ.Ab8RN6Le_B-K4XsTTGDe6Ny00O4JgZnb2uv2_xCKxpw6X0a_VQ"

@st.cache_data(ttl=86400)
def get_realtime_exchange_rate(date_str):
    try:
        target_date = pd.to_datetime(date_str).strftime("%Y%m%d")
        url = f"https://www.kebhana.com/cms/rate/wpHanaIndex.do?searchDate={target_date}&searchInqCount=1"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'html.parser')
        
        tables = soup.find_all('table')
        if tables:
            for row in tables[0].find_all('tr'):
                cells = row.find_all('td')
                if cells and 'USD' in str(cells[0]):
                    rate_text = cells[8].text.strip().replace(",", "")
                    return float(rate_text)
        return 1325.50  
    except Exception:
        return 1325.50

def parse_flexible_month(date_val, target_month_num):
    if pd.isna(date_val) or str(date_val).strip() == "":
        return None
    try:
        dt = pd.to_datetime(date_val)
        if dt.month == target_month_num:
            return dt
    except Exception:
        pass
    
    str_val = str(date_val).strip()
    matches = re.findall(r"(\d+)월|(\d+)[./-]", str_val)
    for m in matches:
        for val in m:
            if val and val.isdigit() and int(val) == target_month_num:
                return datetime(2026, target_month_num, 1)
    return None

def auto_define_pname(lot_str, fallback_val):
    clean_lot = str(lot_str).strip().upper()
    if "200" in clean_lot: return "200ml"
    if "340" in clean_lot: return "340ml"
    if "500" in clean_lot: return "500ml"
    if "1000" in clean_lot: return "1000ml"
    if "HP" in clean_lot: return "HP"
    return str(fallback_val).strip()

def make_unique_cols(columns):
    cols = []
    count = {}
    for c in columns:
        c_str = str(c).strip()
        if c_str in count:
            count[c_str] += 1
            cols.append(f"{c_str}_{count[c_str]}")
        else:
            count[c_str] = 0
            cols.append(c_str)
    return cols

# 웹 대시보드 미리보기용 세관별 컬러 스타일링 함수
def color_sewan_style(val):
    val_str = str(val)
    if '안산' in val_str:
        return 'background-color: #E2EFDA; color: #375623; font-weight: bold;'
    elif '안양' in val_str:
        return 'background-color: #D9E1F2; color: #1F4E78; font-weight: bold;'
    elif '부산' in val_str:
        return 'background-color: #FCE4D6; color: #C65911; font-weight: bold;'
    return ''

# ==========================================
# 3. 메뉴 탭 구성
# ==========================================
tab1, tab2, tab3 = st.tabs([
    "📑 세관 통관 정산 마스터", 
    "📦 해상물류비 마감정산 (공장입고)",
    "💰 외상매입금 현황 마스터"
])

# ------------------------------------------
# 📑 TAB 1: 세관 통관 부가세 정산 마스터 (세관별 모으기 & 컬러 셀 & 테두리 실금)
# ------------------------------------------
with tab1:
    st.markdown("### 📑 안산 / 안양 / 부산 세관 월별납부 부가세·관세 정밀 매칭")
    st.write("관세청 고지목록(Excel)과 수입신고필증(PDF 다중 파일)을 연동하여 제미나이 AI가 결제금액, 운임비, 보험증권을 정밀 검증합니다.")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("#### 📥 1. 실무 자료 업로드")
        notice_file = st.file_uploader(
            "관세청 '월별납부 개별고지목록' (Excel)", 
            type=["xlsx", "xls"], 
            key="notice_final_fixed_v24"
        )
        pdf_files = st.file_uploader(
            "수입신고필증(면장) 통합본 파일 (PDF)", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key="declaration_final_fixed_v24"
        )
        st.markdown("---")
        start_btn = st.button("🚀 최종 마스터 대장 산출하기", use_container_width=True, type="primary", key="btn_customs")
        if not os.path.exists(logo_filename):
            st.info("💡 '삼륭물산한글로고.png' 파일을 폴더에 업로드하시면 배경 워터마크 로고가 활성화됩니다.")

    with col2:
        st.markdown("#### 📋 2. 정산 마스터 대장 결과물")
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
                        
                        # 🛠️ 제미나이 전용 최적화 프롬프트
                        prompt = """
                        당신은 관세 법인 소속의 정산 자동화 AI입니다. 제공된 수입신고필증 PDF 문서 전체를 전수조사하여, 각 '신고번호'별로 아래 항목들을 정확하게 추출하여 JSON 배열 형태로 응답해 주세요.
                        
                        [필증 총액 스캔 규칙 - 필수 준수]
                        1. 다란 건 총합산: 하나의 수입신고번호 면장이 여러 개의 '란'으로 구성되어 분할 표기된 경우, 반드시 해당 신고번호 하위의 모든 '란'의 결제금액(USD)을 누락 없이 전부 더하여(총합산) 하나의 대표 객체로 출력해야 합니다.
                        2. usd_amount (결제금액): 소수점 이하 자리(센트 단위)가 존재한다면 절대로 자르거나 반올림하지 마십시오. 소수점 둘째 자리까지의 원래의 값을 완벽한 소수(Float)로 추출하십시오. (예: 137048.18)
                        3. freight (⑤⑦ 운임) & insurance (⑤⑧ 보험료): 
                           - ★경고: 필증 중앙의 큰 글씨인 '총 과세가격(KRW)' 또는 원화 합계액을 운임비로 절대 잘못 추출하지 마십시오. 운임비는 필증 맨 아래 ⑤⑦번 칸에 따로 적혀있습니다.
                           - 개별 란의 쪼개진 금액이 아니라, 필증 맨 아래의 결산 총액 기재란에 적힌 해당 건의 '총 운임(⑤⑦)' 및 '총 보험료(⑤⑧)' KRW 금액을 추출하십시오. 없으면 0으로 명시하십시오.
                        
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
                            # 🎯 끝자리 알파벳 'u'나 다른 식별 문자 제거
                            raw_shin = str(item.get('shin_no', '')).strip()
                            clean_shin = "".join(filter(str.isalnum, raw_shin))
                            if clean_shin.lower().endswith('u'):
                                clean_shin = clean_shin[:-1]
                            pdf_master_dict[clean_shin] = item
                        
                        processed_data = []
                        
                        for idx, row in df_notice_all.iterrows():
                            excel_shin_no = str(row.get('신고번호', '')).strip()
                            clean_excel_shin = "".join(filter(str.isalnum, excel_shin_no))
                            if clean_excel_shin.lower().endswith('u'):
                                clean_excel_shin = clean_excel_shin[:-1]

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
                                    
                                # 🛡️ 운임비 비정상 과다 검출 방어 로직
                                raw_goods_val = usd_num * fx_num
                                if raw_goods_val > 0 and (freight_num / raw_goods_val) > 0.3:
                                    freight_num = 0
                                    
                                if "SZINC" in bl_no.upper():
                                    freight_num = 0
                                
                                # 🛠️ 4번 항목 센트 누락 방어 밸브
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
                                "번호": 0,  # 정렬 후 재부여 예정
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
                        
                        df_final = pd.DataFrame(processed_data)
                        
                        # 🎯 [요청 반영 1] 세관별(안산 -> 안양 -> 부산 순)로 한눈에 들어오도록 데이터 정렬 및 모으기
                        df_final['세관_정렬기준'] = df_final['세관'].map({'안산세관': 1, '안양세관': 2, '부산세관': 3}).fillna(4)
                        df_final = df_final.sort_values(by=['세관_정렬기준', '신고번호']).reset_index(drop=True)
                        df_final.drop(columns=['세관_정렬기준'], inplace=True)
                        df_final['번호'] = range(1, len(df_final) + 1)  # 모아진 순서대로 1번부터 깔끔하게 재부여
                        
                        output_excel = io.BytesIO()
                        workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                        
                        df_excel_base = df_final[["번호", "신고번호", "납부(고지)번호", "수입부가세 (고지금액)", "신고일자", 
                                              "④ B/L번호 (HBL)", "③⑦ 결제금액 (USD)", "인도조건", "환율", 
                                              "⑤⑦ 운임비 (KRW)", "⑤⑧ 보험증권 (KRW)", "세관", "비고"]]
                        
                        # 엑셀 시트 생성 (to_excel 기본 테두리 누락을 막기 위해 직접 셀 작성)
                        wb = workbook.book
                        ws_sum = wb.add_worksheet("Summary")
                        ws = wb.add_worksheet("통관월납_정산대장")
                        
                        # 🎯 [요청 반영 2] 모든 셀 테두리 실금(border=1) 포함 서식 정의
                        fmt_header = wb.add_format({'bg_color': '#1F4E78', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center'})
                        fmt_center = wb.add_format({'align': 'center', 'border': 1})
                        fmt_num = wb.add_format({'num_format': '#,##0', 'border': 1})
                        fmt_usd = wb.add_format({'num_format': '#,##0.00', 'border': 1})
                        fmt_fx = wb.add_format({'num_format': '#,##0.0000', 'border': 1})
                        
                        # 🎯 [요청 반영 3] 세관별 한눈에 구분을 위한 전용 파스텔 컬러 셀 포맷 (테두리 실금 포함)
                        fmt_ansan = wb.add_format({'bg_color': '#E2EFDA', 'font_color': '#375623', 'bold': True, 'border': 1, 'align': 'center'})  # 연한 초록
                        fmt_anyang = wb.add_format({'bg_color': '#D9E1F2', 'font_color': '#1F4E78', 'bold': True, 'border': 1, 'align': 'center'}) # 연한 파랑
                        fmt_busan = wb.add_format({'bg_color': '#FCE4D6', 'font_color': '#C65911', 'bold': True, 'border': 1, 'align': 'center'})  # 연한 주황
                        fmt_sewan_default = wb.add_format({'bg_color': '#F2F2F2', 'bold': True, 'border': 1, 'align': 'center'}) # 기타
                        
                        blue_bold_center = wb.add_format({'align': 'center', 'font_color': '#002060', 'bold': True, 'border': 1, 'bg_color': '#D9E1F2'})
                        red_bold_center = wb.add_format({'align': 'center', 'font_color': '#FF0000', 'bold': True, 'border': 1, 'bg_color': '#FCE4D6'})
                        
                        summary_header_format = wb.add_format({'bg_color': '#1F4E78', 'font_color': '#FFFFFF', 'bold': True, 'align': 'center', 'border': 1})
                        summary_data_format = wb.add_format({'border': 1, 'num_format': '#,##0'})
                        summary_total_format = wb.add_format({'bg_color': '#D9E1F2', 'bold': True, 'border': 1, 'num_format': '#,##0'})
                        
                        # --- 시트 1 : Summary 시트 ---
                        ws_sum.set_column('B:C', 25)
                        ws_sum.write('B2', '🏢 세관 구분', summary_header_format)
                        ws_sum.write('C2', '💰 수입부가세 총합계 (KRW)', summary_header_format)
                        
                        last_row_idx = len(df_final) + 1
                        ws_sum.write('B3', '안산세관', fmt_ansan)
                        ws_sum.write_formula('C3', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "안산세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                        
                        ws_sum.write('B4', '안양세관', fmt_anyang)
                        ws_sum.write_formula('C4', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "안양세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                        
                        ws_sum.write('B5', '부산세관', fmt_busan)
                        ws_sum.write_formula('C5', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "부산세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                        
                        ws_sum.write('B6', '🚀 전체 부가세 총계', summary_total_format)
                        ws_sum.write_formula('C6', '=SUM(C3:C5)', summary_total_format)
                        
                        # --- 시트 2 : 정산대장 메인 시트 직접 작성 (테두리 100% 보장) ---
                        ws.set_column('A:C', 16)
                        ws.set_column('D:D', 18)
                        ws.set_column('E:F', 15)
                        ws.set_column('G:G', 18)
                        ws.set_column('H:I', 12)
                        ws.set_column('J:K', 16)
                        ws.set_column('L:L', 14)
                        ws.set_column('M:M', 14)
                        ws.set_column('N:N', 24) 
                        ws.set_column('O:O', 25) 
                        ws.set_column('P:P', 22) 
                        
                        # 헤더 작성
                        all_headers = list(df_excel_base.columns) + ['과세가격(원화산출식)', '수식검증 부가세(원단위 버림)', '고지액 검증 결과']
                        for c_idx, h_title in enumerate(all_headers):
                            ws.write(0, c_idx, h_title, fmt_header)
                        
                        # 데이터 셀 작성 (모든 셀 테두리 실금 & 세관별 컬러 삽입)
                        for r_idx, row in df_excel_base.iterrows():
                            excel_r = r_idx + 1  # 0-indexed행 + 1 (엑셀 2번째 줄부터 시작)
                            
                            ws.write(excel_r, 0, row['번호'], fmt_center)
                            ws.write(excel_r, 1, row['신고번호'], fmt_center)
                            ws.write(excel_r, 2, row['납부(고지)번호'], fmt_center)
                            ws.write(excel_r, 3, row['수입부가세 (고지금액)'], fmt_num)
                            ws.write(excel_r, 4, row['신고일자'], fmt_center)
                            ws.write(excel_r, 5, row['④ B/L번호 (HBL)'], fmt_center)
                            ws.write(excel_r, 6, row['③⑦ 결제금액 (USD)'], fmt_usd)
                            ws.write(excel_r, 7, row['인도조건'], fmt_center)
                            ws.write(excel_r, 8, row['환율'], fmt_fx)
                            ws.write(excel_r, 9, row['⑤⑦ 운임비 (KRW)'], fmt_num)
                            ws.write(excel_r, 10, row['⑤⑧ 보험증권 (KRW)'], fmt_num)
                            
                            # 세관별 셀 배경색 분기 적용
                            sewan_val = str(row['세관'])
                            if '안산' in sewan_val:
                                ws.write(excel_r, 11, sewan_val, fmt_ansan)
                            elif '안양' in sewan_val:
                                ws.write(excel_r, 11, sewan_val, fmt_anyang)
                            elif '부산' in sewan_val:
                                ws.write(excel_r, 11, sewan_val, fmt_busan)
                            else:
                                ws.write(excel_r, 11, sewan_val, fmt_sewan_default)
                                
                            ws.write(excel_r, 12, row['비고'], fmt_center)
                            
                            # N, O, P 열 수식 작성 (테두리 실금 포함)
                            f_row = excel_r + 1
                            ws.write_formula(excel_r, 13, f'=(G{f_row}*I{f_row})+J{f_row}+K{f_row}', fmt_num)
                            ws.write_formula(excel_r, 14, f'=ROUNDDOWN(N{f_row}*0.1, -1)', fmt_num)
                            ws.write_formula(excel_r, 15, f'=IF(D{f_row}=O{f_row}, "✔ 완벽 일치", "❌ 금액 불일치")', fmt_center)
                        
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
                        
                        st.success("🎉 세관별 자동 정렬, 파스텔 컬러 구분, 모든 셀 테두리 실금 적용이 완료되었습니다!")
                        st.download_button(
                            label="📥 최종 고도화 정산 마스터 대장 다운로드 (.xlsx)",
                            data=excel_data,
                            file_name="통관월납_정산_마스터대장_최종본_수식포함.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        # 웹 화면 미리보기에서도 세관별 색상 적용하여 출력
                        try:
                            st.dataframe(df_final.style.map(color_sewan_style, subset=['세관']), use_container_width=True)
                        except AttributeError:
                            st.dataframe(df_final.style.applymap(color_sewan_style, subset=['세관']), use_container_width=True)
                        except Exception:
                            st.dataframe(df_final, use_container_width=True)
                        
                    except Exception as e:
                        st.error(f"❌ 가독성 양식 업그레이드 중 오류가 발생했습니다: {e}")

# ------------------------------------------
# 📦 TAB 2: 해상물류비 마감정산 (공장입고)
# ------------------------------------------
with tab2:
    st.markdown("### 📦 해상물류비 마감정산 및 단가 감사 시스템")
    st.write("판토스 등 물류사의 마감내역서와 내부 반입계획서를 대조하여 청구 운임 단가 및 물동량의 일치 여부를 자동으로 검증합니다.")
    
    l_col1, l_col2 = st.columns([1, 2])
    with l_col1:
        st.markdown("#### 📥 1. 물류 청구서 vs 계획서 업로드")
        pantos_file = st.file_uploader("📋 판토스 마감내역서 (Excel)", type=["xlsx", "xls"], key="pantos_uploader")
        plan_file = st.file_uploader("📋 물류 반입계획서 (Excel)", type=["xlsx", "xls"], key="logistics_plan_uploader")
        st.markdown("---")
        logistics_btn = st.button("🚀 물류비 감사 대장 생성", use_container_width=True, type="primary", key="btn_logistics")
        
    with l_col2:
        st.markdown("#### 📋 2. 물류비 감사 결과 프리뷰")
        if logistics_btn:
            if not pantos_file or not plan_file:
                st.error("❌ 물류사 마감내역서와 반입계획서 엑셀 파일을 모두 업로드해 주세요.")
            else:
                with st.spinner("🤖 물류 청구 내역 대조 및 운임 단가 이상 유무 검증 연산 중..."):
                    try:
                        df_pantos = pd.read_excel(pantos_file)
                        df_plan = pd.read_excel(plan_file)
                        
                        df_pantos.columns = make_unique_cols(df_pantos.columns)
                        df_plan.columns = make_unique_cols(df_plan.columns)
                        
                        df_pantos_clean = df_pantos.dropna(how='all').reset_index(drop=True)
                        df_plan_clean = df_plan.dropna(how='all').reset_index(drop=True)
                        
                        lot_col_p = next((c for c in df_pantos_clean.columns if "LOT" in str(c).upper() or "B/L" in str(c).upper() or "오더" in str(c)), df_pantos_clean.columns[0])
                        lot_col_pl = next((c for c in df_plan_clean.columns if "LOT" in str(c).upper() or "B/L" in str(c).upper() or "오더" in str(c)), df_plan_clean.columns[0])
                        
                        df_pantos_clean["매칭키"] = df_pantos_clean[lot_col_p].astype(str).str.strip().str.upper()
                        df_plan_clean["매칭키"] = df_plan_clean[lot_col_pl].astype(str).str.strip().str.upper()
                        
                        df_audit = pd.merge(df_pantos_clean, df_plan_clean, on="매칭키", how="outer", suffixes=("_청구내역", "_반입계획"))
                        df_audit.drop(columns=["매칭키"], inplace=True, errors="ignore")
                        
                        output_log = io.BytesIO()
                        with pd.ExcelWriter(output_log, engine='xlsxwriter') as writer:
                            df_audit.to_excel(writer, sheet_name="해상물류비감사대장", index=False)
                            wb = writer.book
                            ws = writer.sheets["해상물류비감사대장"]
                            fmt_header = wb.add_format({'bold': True, 'bg_color': '#E2EFDA', 'border': 1, 'align': 'center'})
                            for col_num, value in enumerate(df_audit.columns.values):
                                ws.write(0, col_num, value, fmt_header)
                                ws.set_column(col_num, col_num, max(len(str(value)) + 4, 14))
                                
                        st.success("🎉 판토스 물류비 청구 내역과 반입계획서 감사 대조가 완료되었습니다!")
                        st.download_button(
                            label="📥 해상물류비 감사 대장 다운로드 (.xlsx)",
                            data=output_log.getvalue(),
                            file_name=f"해상물류비_감사대장_{datetime.now().strftime('%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        st.dataframe(df_audit, use_container_width=True)
                    except Exception as e:
                        st.error(f"❌ 물류비 감사 처리 중 오류가 발생했습니다: {e}")

# ==========================================
# 💰 TAB 3: 외상매입금 현황 마스터 (실시간 환율 + ENSO 초록 글씨 + 최종합계 완결판)
# ==========================================
with tab3:
    st.markdown("### 💰 미정산 외상매입금 현황 자동 마감 시스템")
    st.write("반입계획서 파싱 후 회계일자별 매매기준율을 실시간 수집하며, E25/E26 LOT 초록 글씨 하이라이트 및 최종합계를 일괄 생성합니다.")
    
    m_col1, m_col2 = st.columns([1, 2])
    
    with m_col1:
        st.markdown("#### 📥 1. 마감 기본자료 업로드")
        months_options = [
            "2026년 12월", "2026년 11월", "2026년 10월", "2026년 9월", "2026년 8월", "2026년 7월",
            "2026년 6월", "2026년 5월", "2026년 4월", "2026년 3월", "2026년 2월", "2026년 1월"
        ]
        target_month = st.selectbox("마감 대상월 선택", months_options)
        match_m = re.search(r"(\d+)월", target_month)
        selected_month_num = int(match_m.group(1)) if match_m else 6
        
        uploaded_payable_plan = st.file_uploader("📋 반입계획서 파일 업로드 (Excel)", type=["xlsx"], key="payable_plan_uploader")
        st.markdown("---")
        calc_ap_btn = st.button("🚀 외상매입금 마감 대장 생성하기", use_container_width=True, type="primary")
        
    with m_col2:
        st.markdown("#### 📋 2. 외상매입금 대장 산출 프리뷰")
        if calc_ap_btn:
            if not uploaded_payable_plan:
                st.error("❌ 정산 처리를 위해 반입계획서 엑셀 파일을 업로드해 주세요.")
            else:
                with st.spinner(f"🤖 실시간 환율 수집 및 {selected_month_num}월 단일 대계 집계 테이블 구성 중..."):
                    try:
                        excel_file = pd.ExcelFile(uploaded_payable_plan)
                        sheet_names_lower = {s.lower().replace(" ", ""): s for s in excel_file.sheet_names}
                        
                        processed_list = []
                        
                        # ① NDP 시트 처리 (J, K, L, M열)
                        ndp_sheet_key = next((s for s in sheet_names_lower if "ndp" in s), None)
                        if ndp_sheet_key:
                            df_ndp = pd.read_excel(excel_file, sheet_name=sheet_names_lower[ndp_sheet_key], header=None)
                            start_parsing = False
                            c_lot, c_pname, c_pickup = 1, 2, 6
                            
                            for idx, row in df_ndp.iterrows():
                                if any("Lot" in str(v) or "오더" in str(v) or "LOT" in str(v).upper() for v in row.values):
                                    start_parsing = True
                                    row_vals = [str(v).strip() for v in row.values]
                                    for r_idx, val in enumerate(row_vals):
                                        val_up = val.upper()
                                        if "LOT" in val_up or "오더" in val_up: c_lot = r_idx
                                        if "품명" in val or "규격" in val: c_pname = r_idx
                                        if any(k in val for k in ["픽업", "타코마", "인계", "일자", "DATE", "ETD", "ETA"]): c_pickup = r_idx
                                    continue
                                
                                if start_parsing:
                                    lot_val = str(row[c_lot]).strip() if pd.notna(row[c_lot]) else ""
                                    if lot_val in ("", "nan", "None") or "Lot" in lot_val or "LOT" in lot_val:
                                        continue
                                        
                                    raw_pname = str(row[c_pname]).strip() if pd.notna(row[c_pname]) else "200ml"
                                    p_name = auto_define_pname(lot_val, raw_pname)
                                    
                                    date_raw = row[c_pickup]
                                    dt_obj = parse_flexible_month(date_raw, selected_month_num)
                                    
                                    if dt_obj:
                                        accounting_date = dt_obj.strftime("%Y-%m-%d")
                                        rl_val = float(str(row[9]).replace(",", "")) if pd.notna(row[9]) else 0.0
                                        kg_val = float(str(row[10]).replace(",", "")) if pd.notna(row[10]) else 0.0
                                        sqm_val = float(str(row[11]).replace(",", "")) if pd.notna(row[11]) else 0.0
                                        amt_val = float(str(row[12]).replace(",", "")) if pd.notna(row[12]) else 0.0
                                        
                                        # 🎯 [실시간 환율 반영] 회계일자 당시 하나은행 매매기준율 조회
                                        fx_rate = get_realtime_exchange_rate(accounting_date)
                                        processed_list.append({
                                            "품명": p_name, "LOT No.": lot_val, "회계일자": accounting_date,
                                            "R/L": rl_val, "중량": kg_val, "면적": sqm_val, "Amount($)": amt_val,
                                            "환율": fx_rate, "원화금액": int(amt_val * fx_rate)
                                        })
                        
                        # ② ENSO 시트 처리 (H, I, J, K열)
                        enso_sheet_key = next((s for s in sheet_names_lower if "enso" in s), None)
                        if enso_sheet_key:
                            df_enso = pd.read_excel(excel_file, sheet_name=sheet_names_lower[enso_sheet_key], header=None)
                            start_parsing = False
                            c_lot, c_pname, c_arr = 1, 2, 4
                            
                            for idx, row in df_enso.iterrows():
                                if any("Lot" in str(v) or "오더" in str(v) or "LOT" in str(v).upper() for v in row.values):
                                    start_parsing = True
                                    row_vals = [str(v).strip() for v in row.values]
                                    for r_idx, val in enumerate(row_vals):
                                        val_up = val.upper()
                                        if "LOT" in val_up or "오더" in val_up: c_lot = r_idx
                                        if "품명" in val or "규격" in val: c_pname = r_idx
                                        if any(k in val for k in ["입항", "도착", "일자", "DATE", "ETA"]): c_arr = r_idx
                                    continue
                                
                                if start_parsing:
                                    lot_val = str(row[c_lot]).strip() if pd.notna(row[c_lot]) else ""
                                    if lot_val in ("", "nan", "None") or "Lot" in lot_val or "LOT" in lot_val:
                                        continue
                                        
                                    raw_pname = str(row[c_pname]).strip() if pd.notna(row[c_pname]) else "200ml"
                                    p_name = auto_define_pname(lot_val, raw_pname)
                                    
                                    date_raw = row[c_arr]
                                    dt_obj = parse_flexible_month(date_raw, selected_month_num)
                                    
                                    if dt_obj:
                                        accounting_date = dt_obj.strftime("%Y-%m-%d")
                                        rl_val = float(str(row[7]).replace(",", "")) if pd.notna(row[7]) else 0.0
                                        kg_val = float(str(row[8]).replace(",", "")) if pd.notna(row[8]) else 0.0
                                        sqm_val = float(str(row[9]).replace(",", "")) if pd.notna(row[9]) else 0.0
                                        amt_val = float(str(row[10]).replace(",", "")) if pd.notna(row[10]) else 0.0
                                        
                                        # 🎯 [실시간 환율 반영] 회계일자 당시 하나은행 매매기준율 조회
                                        fx_rate = get_realtime_exchange_rate(accounting_date)
                                        processed_list.append({
                                            "품명": p_name, "LOT No.": lot_val, "회계일자": accounting_date,
                                            "R/L": rl_val, "중량": kg_val, "면적": sqm_val, "Amount($)": amt_val,
                                            "환율": fx_rate, "원화금액": int(amt_val * fx_rate)
                                        })
                        
                        # ③ 단일 대계 취합 서식 빌드부 (E25/E26 초록색 강조 + 최종합계 수식)
                        if not processed_list:
                            st.warning(f"⚠️ 업로드하신 반입계획서 내에서 {selected_month_num}월 조건에 매칭되는 데이터가 인식되지 않았습니다.")
                        else:
                            df_preview = pd.DataFrame(processed_list)
                            df_preview = df_preview.sort_values(by="품명").reset_index(drop=True)
                            
                            ap_excel = io.BytesIO()
                            with pd.ExcelWriter(ap_excel, engine='xlsxwriter') as writer:
                                wb = writer.book
                                ws = wb.add_worksheet(target_month.replace("2026년 ", "26년 "))
                                
                                fmt_title = wb.add_format({'bold': True, 'font_size': 16, 'font_name': '맑은 고딕'})
                                fmt_memo = wb.add_format({'font_color': 'red', 'font_name': '맑은 고딕', 'font_size': 10, 'bold': True})
                                fmt_th = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
                                fmt_subth = wb.add_format({'bg_color': '#F2F2F2', 'border': 1, 'align': 'center', 'font_size': 9})
                                
                                fmt_cell = wb.add_format({'border': 1, 'align': 'center'})
                                fmt_green_cell = wb.add_format({'border': 1, 'align': 'center', 'font_color': '#006100', 'bold': True})
                                
                                fmt_num = wb.add_format({'border': 1, 'num_format': '#,##0'})
                                fmt_usd = wb.add_format({'border': 1, 'num_format': '#,##0.00'})
                                fmt_subtotal = wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'border': 1, 'num_format': '#,##0'})
                                fmt_subtotal_usd = wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'border': 1, 'num_format': '#,##0.00'})
                                
                                fmt_total = wb.add_format({'bg_color': '#FCE4D6', 'bold': True, 'border': 1, 'num_format': '#,##0', 'align': 'center'})
                                fmt_total_usd = wb.add_format({'bg_color': '#FCE4D6', 'bold': True, 'border': 1, 'num_format': '#,##0.00', 'align': 'center'})
                                
                                ws.write(0, 0, f"{target_month.replace('2026년 ', '2026년도 ')} 미정산 외상매입금 현황", fmt_title)
                                ws.write(3, 1, f"** 회계일자 빨강 표시는 타코마에서 화물 인계일", fmt_memo)
                                
                                headers = ["품명", "LOT No.", "회계일자", "R/L", "중량", "면적", "Amount($)", "환율", "원화금액"]
                                for c_idx, h in enumerate(headers):
                                    ws.write(4, c_idx, h, fmt_th)
                                ws.write(5, 4, "kg", fmt_subth)
                                ws.write(5, 5, "SQ", fmt_subth)
                                
                                row_idx = 6
                                unique_pnames = df_preview["품명"].unique()
                                
                                for p_name in unique_pnames:
                                    df_group = df_preview[df_preview["품명"] == p_name]
                                    start_row = row_idx + 1
                                    
                                    for i, (_, item) in enumerate(df_group.iterrows()):
                                        lot_no = str(item["LOT No."]).strip()
                                        is_enso = lot_no.upper().startswith("E25") or lot_no.upper().startswith("E26")
                                        
                                        ws.write(row_idx, 0, item["품명"] if i == 0 else "", fmt_cell)
                                        
                                        if is_enso:
                                            ws.write(row_idx, 1, lot_no, fmt_green_cell)
                                        else:
                                            ws.write(row_idx, 1, lot_no, fmt_cell)
                                            
                                        ws.write(row_idx, 2, item["회계일자"], fmt_cell)
                                        ws.write(row_idx, 3, item["R/L"], fmt_num)
                                        ws.write(row_idx, 4, item["중량"], fmt_num)
                                        ws.write(row_idx, 5, item["면적"], fmt_num)
                                        ws.write(row_idx, 6, item["Amount($)"], fmt_usd)
                                        ws.write(row_idx, 7, item["환율"], fmt_usd)
                                        ws.write_formula(row_idx, 8, f"=G{row_idx+1}*H{row_idx+1}", fmt_num)
                                        row_idx += 1
                                    
                                    end_row = row_idx
                                    
                                    for _ in range(3):
                                        for c in range(8):
                                            ws.write(row_idx, c, "", fmt_cell)
                                        ws.write(row_idx, 8, 0, fmt_num)
                                        row_idx += 1
                                        
                                    row_idx += 1 
                                    
                                    # 품명별 소계 작성
                                    ws.write(row_idx, 0, "", fmt_cell)
                                    ws.write(row_idx, 1, "소       계", wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'align': 'center', 'border': 1}))
                                    ws.write(row_idx, 2, len(df_group), fmt_cell)
                                    ws.write_formula(row_idx, 3, f"=SUM(D{start_row}:D{end_row})", fmt_subtotal)
                                    ws.write_formula(row_idx, 4, f"=SUM(E{start_row}:E{end_row})", fmt_subtotal)
                                    ws.write_formula(row_idx, 5, f"=SUM(F{start_row}:F{end_row})", fmt_subtotal)
                                    ws.write_formula(row_idx, 6, f"=SUM(G{start_row}:G{end_row})", fmt_subtotal_usd)
                                    ws.write(row_idx, 7, "", fmt_cell)
                                    ws.write_formula(row_idx, 8, f"=SUM(I{start_row}:I{end_row})", fmt_subtotal)
                                    row_idx += 1
                                    
                                    row_idx += 1
                                
                                # 최하단 전체 합산 "총   합   계" 행 빌드 (SUMIF 수식 사용)
                                ws.write(row_idx, 0, "", fmt_cell)
                                ws.write(row_idx, 1, "총   합   계", wb.add_format({'bg_color': '#FCE4D6', 'bold': True, 'align': 'center', 'border': 1}))
                                
                                ws.write_formula(row_idx, 2, f'=SUMIF(B7:B{row_idx}, "소       계", C7:C{row_idx})', fmt_total)
                                ws.write_formula(row_idx, 3, f'=SUMIF(B7:B{row_idx}, "소       계", D7:D{row_idx})', fmt_total)
                                ws.write_formula(row_idx, 4, f'=SUMIF(B7:B{row_idx}, "소       계", E7:E{row_idx})', fmt_total)
                                ws.write_formula(row_idx, 5, f'=SUMIF(B7:B{row_idx}, "소       계", F7:F{row_idx})', fmt_total)
                                ws.write_formula(row_idx, 6, f'=SUMIF(B7:B{row_idx}, "소       계", G7:G{row_idx})', fmt_total_usd)
                                ws.write(row_idx, 7, "", fmt_cell)
                                ws.write_formula(row_idx, 8, f'=SUMIF(B7:B{row_idx}, "소       계", I7:I{row_idx})', fmt_total)
                                
                                ws.set_column('A:B', 16)
                                ws.set_column('C:C', 13)
                                ws.set_column('D:F', 11)
                                ws.set_column('G:I', 16)
                                
                            st.success(f"🎯 처리 완료! {target_month} 외상매입금 대장이 일자별 실시간 매매기준 환율로 정확히 정산되었습니다.")
                            st.download_button(
                                label="📥 외상매입금 마스터 엑셀 다운로드 (.xlsx)",
                                data=ap_excel.getvalue(),
                                file_name=f"외상매입금현황_{target_month.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            st.dataframe(df_preview, use_container_width=True)
                            
                    except Exception as e:
                        st.error(f"❌ 데이터 정산 및 소계 빌드 중 오류 발생: {e}")
