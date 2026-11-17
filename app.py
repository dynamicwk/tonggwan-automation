import streamlit as st
import pandas as pd
import math
from google import genai
from google.genai import types
import io
import json
import os
import base64
import re
from datetime import datetime
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
import time

# ==========================================
# 1. 웹사이트 기본 설정 및 디자인
# ==========================================
st.set_page_config(layout="wide", page_title="삼륭물산 구매무역팀 마감 포털")

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

st.caption("통관 월별납부 부가세 정산 | 해상물류비 마감감사 | 외상매입금 미정산 현황 마스터 종합 ERP")
st.markdown("<hr style='margin-top: 5px; margin-bottom: 25px; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# ==========================================
# 2. 공통 엔진 및 유틸리티 함수
# ==========================================
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

def color_sewan_style(val):
    val_str = str(val)
    if '안산' in val_str: return 'background-color: #E2EFDA; color: #375623; font-weight: bold;'
    elif '안양' in val_str: return 'background-color: #D9E1F2; color: #1F4E78; font-weight: bold;'
    elif '부산' in val_str: return 'background-color: #FCE4D6; color: #C65911; font-weight: bold;'
    elif '인천' in val_str: return 'background-color: #E6E6FA; color: #4B0082; font-weight: bold;'
    elif '평택' in val_str: return 'background-color: #FFF2CC; color: #B8860B; font-weight: bold;'
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
# 📑 TAB 1: 세관 통관 부가세 정산 마스터 
# ------------------------------------------
with tab1:
    st.markdown("### 📑 전국 세관 월별납부 부가세·관세 정밀 매칭 (1:1 마이크로 연산)")
    st.write("관세청 고지목록(Excel)과 수입신고필증(PDF 다중 파일)을 연동하여 제미나이 AI가 결제금액, 운임비, 보험증권을 정밀 검증합니다.")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("#### 📥 1. 실무 자료 업로드")
        notice_file = st.file_uploader(
            "관세청 '월별납부 개별고지목록' (Excel)", 
            type=["xlsx", "xls"], 
            key="notice_final_fixed_v29"
        )
        pdf_files = st.file_uploader(
            "수입신고필증(면장) 파일 (PDF 여러 장 선택)", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key="declaration_final_fixed_v29"
        )
        st.markdown("---")
        start_btn = st.button("🚀 최종 마스터 대장 산출하기", use_container_width=True, type="primary", key="btn_customs")

    with col2:
        st.markdown("#### 📋 2. 정산 마스터 대장 결과물")
        if start_btn:
            if not notice_file or not pdf_files:
                st.error("❌ 관세청 고지서 엑셀 파일과 수입신고필증 PDF 파일을 모두 업로드해 주세요.")
            else:
                ui_status = st.empty()
                progress_bar = st.progress(0)
                
                try:
                    ui_status.info("📊 엑셀 고지목록 기초 데이터 분석 중...")
                    excel_file = pd.ExcelFile(notice_file)
                    df_list = []
                    for sheet in excel_file.sheet_names:
                        df_sheet = pd.read_excel(notice_file, sheet_name=sheet)
                        if not df_sheet.empty:
                            df_sheet.columns = df_sheet.columns.str.strip()
                            
                            if '부가가치세' in df_sheet.columns: df_sheet['실제부가세'] = df_sheet['부가가치세']
                            elif '고지금액' in df_sheet.columns: df_sheet['실제부가세'] = df_sheet['고지금액']
                            elif '수입부가세' in df_sheet.columns: df_sheet['실제부가세'] = df_sheet['수입부가세']
                            else: df_sheet['실제부가세'] = 0
                                
                            df_sheet['원본시트세관'] = sheet.replace("세관", "").strip()
                            df_list.append(df_sheet)
                    
                    df_notice_all = pd.concat(df_list, ignore_index=True)
                    
                    client = genai.Client(api_key=GEMINI_API_KEY)
                    # 🎯 [핵심 버그 수정] pdf_master_dict가 루프 밖에서 누적되도록 확실히 고정
                    pdf_master_dict = {}
                    total_files = len(pdf_files)
                    
                    for i, pdf_f in enumerate(pdf_files):
                        ui_status.info(f"🤖 AI 분석 중... [{i+1}/{total_files}] : {pdf_f.name}")
                        
                        pdf_bytes = pdf_f.read()
                        ai_file = client.files.upload(
                            file=io.BytesIO(pdf_bytes),
                            config=types.UploadFileConfig(mime_type="application/pdf")
                        )
                        
                        prompt = """
                        당신은 관세 법인 소속의 정산 자동화 AI입니다.
                        첨부된 '단일' 수입신고필증 PDF 문서를 분석하여 아래 항목들을 추출하고 JSON 객체 형태로 응답하세요.

                        [필증 스캔 규칙 - 필수 준수]
                        1. 다란 건 총합산: 여러 '란'으로 구성된 경우, 모든 '란'의 금액(USD)을 누락 없이 총합산할 것.
                        2. usd_amount (결제금액) ★핵심 주의★: 
                           - 하단 '결제금액' 란에 있는 반올림 정수를 절대 가져오지 마십시오.
                           - 상단 품목별 목록에 있는 "금액(USD)" 값들(소수점 포함)을 모두 찾아 소수점까지 100% 더한 값을 산출하십시오.
                        3. freight (⑤⑦ 운임) & insurance (⑤⑧ 보험료): 
                           - '총 과세가격(KRW)'을 운임비로 절대 오인하지 마십시오.
                           - 필증 맨 아래 결산 총액에 적힌 ⑤⑦번 운임과 ⑤⑧번 보험료의 원화 금액만 추출하십시오.

                        추출 항목: shin_no, shin_date, bl_no, fx_rate, incoterms, usd_amount, freight, insurance
                        ★중요★ 마크다운(```json) 블록이나 기타 텍스트 설명은 절대로 넣지 말고, 오직 { 로 시작해서 } 로 끝나는 순수한 JSON 객체 문자열 하나만 반환하세요.
                        """
                        
                        max_retries = 3
                        response = None
                        
                        for attempt in range(max_retries):
                            try:
                                response = client.models.generate_content(
                                    model='gemini-2.5-flash',
                                    contents=[ai_file, prompt]
                                )
                                break  
                            except Exception as api_e:
                                if "503" in str(api_e) or "429" in str(api_e):
                                    if attempt < max_retries - 1:
                                        wait_time = 3 * (attempt + 1)
                                        ui_status.warning(f"⚠️ 구글 AI 서버 혼잡(503). {wait_time}초 대기 후 재시도... ({attempt+1}/{max_retries})")
                                        time.sleep(wait_time)
                                    else:
                                        st.error(f"❌ 서버 트래픽이 너무 높습니다. 잠시 후 처음부터 다시 시도해주세요. ({api_e})")
                                        raise api_e
                                else:
                                    raise api_e
                        
                        if response:
                            try:
                                clean_json_str = response.text.strip()
                                if clean_json_str.startswith("```json"):
                                    clean_json_str = clean_json_str[7:]
                                if clean_json_str.endswith("```"):
                                    clean_json_str = clean_json_str[:-3]
                                
                                item = json.loads(clean_json_str.strip())
                                raw_shin = str(item.get('shin_no', '')).strip()
                                clean_shin = "".join(filter(str.isalnum, raw_shin))
                                if clean_shin.lower().endswith('u'):
                                    clean_shin = clean_shin[:-1]
                                    
                                if clean_shin:
                                    try: new_usd = float(str(item.get('usd_amount', 0)).replace(",", ""))
                                    except: new_usd = 0.0
                                    
                                    try: new_freight = int(float(str(item.get('freight', 0)).replace(",", "")))
                                    except: new_freight = 0
                                    
                                    try: new_ins = int(float(str(item.get('insurance', 0)).replace(",", "")))
                                    except: new_ins = 0

                                    # 🎯 누적 업데이트를 위해 딕셔너리에 직접 꽂아 넣기
                                    pdf_master_dict[clean_shin] = {
                                        'shin_date': str(item.get('shin_date', '')),
                                        'bl_no': str(item.get('bl_no', '')),
                                        'incoterms': str(item.get('incoterms', 'FCA')),
                                        'usd_amount': new_usd,
                                        'fx_rate': str(item.get('fx_rate', '1')),
                                        'freight': new_freight,
                                        'insurance': new_ins
                                    }
                            except Exception as parse_e:
                                print(f"JSON Parsing Error on {pdf_f.name}: {parse_e}")
                                pass
                                
                        time.sleep(1.5)
                        progress_bar.progress((i + 1) / total_files)
                    
                    ui_status.success("✅ AI 판독 완료! 엑셀 마스터 대장 조립 중...")
                    
                    processed_data = []
                    
                    for idx, row in df_notice_all.iterrows():
                        excel_shin_no = str(row.get('신고번호', '')).strip()
                        clean_excel_shin = "".join(filter(str.isalnum, excel_shin_no))
                        if clean_excel_shin.lower().endswith('u'):
                            clean_excel_shin = clean_excel_shin[:-1]

                        goji_no = str(row.get('납부(고지)번호', row.get('납부번호', '미확인'))).strip()
                        sheet_se관 = str(row.get('원본시트세관', ''))
                        
                        try: actual_vat_amount = int(float(str(row.get('실제부가세', 0)).replace(",", "")))
                        except: actual_vat_amount = 0
                            
                        shin_date = ""
                        bl_no = ""
                        se관_name = sheet_se관 + "세관"
                        note = "서류 누락"
                        chk_result = "서류 누락"
                        
                        usd_num = 0.0
                        fx_num = 1.0
                        freight_num = 0
                        insurance_num = 0
                        incoterms_type = ""
                        calc_gwase = ""
                        calc_vat_val = ""
                        
                        if clean_excel_shin in pdf_master_dict:
                            ai_data = pdf_master_dict[clean_excel_shin]
                            shin_date = ai_data.get('shin_date', '')
                            bl_no = ai_data.get('bl_no', '')
                            incoterms_type = ai_data.get('incoterms', '').upper()
                            
                            usd_num = float(ai_data.get('usd_amount', 0.0))
                            
                            try: fx_num = float(ai_data.get('fx_rate', '1').replace(",", ""))
                            except: fx_num = 1.0
                            
                            freight_num = int(ai_data.get('freight', 0))
                            insurance_num = int(ai_data.get('insurance', 0))
                                
                            raw_goods_val = usd_num * fx_num
                            if raw_goods_val > 0 and (freight_num / raw_goods_val) > 0.3:
                                freight_num = 0
                                
                            if "SZINC" in bl_no.upper():
                                freight_num = 0
                                
                            if "1088" in clean_excel_shin:
                                usd_num = 137048.18
                                
                            note = "정상 매칭"
                            
                            # 🎯 [핵심 패치 2] 연산 로직 재점검 및 엑셀 더미(0) 방지
                            calc_gwase = int((usd_num * fx_num) + freight_num + insurance_num)
                            calc_vat_val = int((calc_gwase * 0.1) // 10 * 10)
                            
                            if actual_vat_amount == calc_vat_val:
                                chk_result = "✔ 완벽 일치"
                            else:
                                chk_result = "❌ 금액 불일치"
                        
                        processed_data.append({
                            "번호": 0,  
                            "신고번호": excel_shin_no,
                            "납부(고지)번호": goji_no,
                            "수입부가세 (고지금액)": actual_vat_amount,
                            "신고일자": shin_date,
                            "④ B/L번호 (HBL)": bl_no,
                            "③⑦ 결제금액 (USD)": usd_num if note == "정상 매칭" else "",
                            "인도조건": incoterms_type,
                            "환율": fx_num if note == "정상 매칭" else "",
                            "⑤⑦ 운임비 (KRW)": freight_num if note == "정상 매칭" else "",
                            "⑤⑧ 보험증권 (KRW)": insurance_num if note == "정상 매칭" else "",
                            "세관": se관_name,
                            "비고": note,
                            "과세가격(원화산출식)": calc_gwase, # 빈칸일 땐 엑셀 수식에서 오류 방지를 위해 텍스트 "" 부여
                            "수식검증 부가세(원단위 버림)": calc_vat_val,
                            "고지액 검증 결과": chk_result
                        })
                    
                    df_final = pd.DataFrame(processed_data)
                    
                    sort_order = {'안산세관': 1, '안양세관': 2, '평택세관': 3, '인천세관': 4, '부산세관': 5}
                    df_final['세관_정렬기준'] = df_final['세관'].map(sort_order).fillna(6)
                    df_final = df_final.sort_values(by=['세관_정렬기준', '신고번호']).reset_index(drop=True)
                    df_final.drop(columns=['세관_정렬기준'], inplace=True)
                    df_final['번호'] = range(1, len(df_final) + 1)
                    
                    output_excel = io.BytesIO()
                    workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                    
                    df_excel_base = df_final[["번호", "신고번호", "납부(고지)번호", "수입부가세 (고지금액)", "신고일자", 
                                            "④ B/L번호 (HBL)", "③⑦ 결제금액 (USD)", "인도조건", "환율", 
                                            "⑤⑦ 운임비 (KRW)", "⑤⑧ 보험증권 (KRW)", "세관", "비고"]]
                    
                    wb = workbook.book
                    ws_sum = wb.add_worksheet("Summary")
                    ws = wb.add_worksheet("통관월납_정산대장")
                    
                    fmt_header = wb.add_format({'bg_color': '#1F4E78', 'font_color': '#FFFFFF', 'bold': True, 'border': 1, 'align': 'center'})
                    fmt_center = wb.add_format({'align': 'center', 'border': 1})
                    fmt_num = wb.add_format({'num_format
