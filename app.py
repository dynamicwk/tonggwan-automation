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

# 웹사이트 설정 및 디자인 (넓은 화면 모드)
st.set_page_config(layout="wide", page_title="삼륭물산 구매무역팀 마감 포털")

# 🖼️ [워터마크 엔진 - 은은한 중앙 배경 복원 모델]
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
        background-size: 35%; 
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
            <span style="font-size: 15px; color: #333333; font-weight: bold; background-color: rgba(255, 255, 255, 0.9); padding: 6px 14px; border-radius: 6px; border: 1px solid #ddd;">
                🏢 삼륭물산 구매무역팀
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.caption("안산 / 안양 / 부산 세관 월별납부 정산 | 해상물류비 마감감사 | 외상매입금 미정산 현황 마스터 종합 ERP")
st.markdown("<hr style='margin-top: 5px; margin-bottom: 25px; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# 🔑 구글 제미나이 API 키 고정
GEMINI_API_KEY = "AQ.Ab8RN6Le_B-K4XsTTGDe6Ny00O4JgZnb2uv2_xCKxpw6X0a_VQ"

# 🗄️ [하나은행 최초 매매기준율 실시간 수집 크롤러 핵심 엔진]
@st.cache_data(ttl=3600)
def get_hana_first_exrate(date_str):
    try:
        clean_date = date_str.replace("-", "")
        url = f"https://www.kebhana.com/cms/rate/wpHanaIndex.do?searchDate={clean_date}&searchInqCount=1"
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

# 비정형 한글 텍스트 날짜 유연 필터링 처리 함수 (연도 예외 극복용 강화 패치)
def parse_flexible_month(date_val, target_month_num):
    if pd.isna(date_val) or str(date_val).strip() == "":
        return None
    try:
        dt = pd.to_datetime(date_val)
        if dt.month == target_month_num:
            return dt
    except Exception:
        pass
    
    # 일반 텍스트로 기록된 날짜(ex. "06.11", "6/11", "8월말 9월중순")에서 월 숫자 강제 추출
    str_val = str(date_val).strip()
    matches = re.findall(r"(\d+)월|(\d+)[./-]", str_val)
    for m in matches:
        for val in m:
            if val and val.isdigit() and int(val) == target_month_num:
                # 타겟 월과 일치하면 임시 날짜 객체 생성하여 통과
                return datetime(2026, target_month_num, 1)
    return None

# 삼륭물산 고유의 LOT 번호 패턴 매칭 기반 품명 강제 자동 정합성 엔진
def auto_define_pname(lot_str, fallback_val):
    clean_lot = str(lot_str).strip().upper()
    if "200" in clean_lot: return "200ml"
    if "340" in clean_lot: return "340ml"
    if "500" in clean_lot: return "500ml"
    if "1000" in clean_lot: return "1000ml"
    if "HP" in clean_lot: return "HP"
    return str(fallback_val).strip()

# 🗂️ 3대 대메뉴 기능 탭 분할
tab1, tab2, tab3 = st.tabs([
    "📑 세관 통관 정산 마스터", 
    "📦 해상물류비 마감정산 (공장입고)",
    "💰 외상매입금 현황 마스터"
])

with tab1:
    st.info("세관 통관 정산 서비스가 대기 중입니다.")
with tab2:
    st.info("해상물류비 정산 서비스가 대기 중입니다.")

# ==========================================
# 💰 탭 3: 외상매입금 현황 마스터 (안정성 강화 버전)
# ==========================================
with tab3:
    st.markdown("### 💰 미정산 외상매입금 현황 자동 마감 시스템")
    st.write("반입계획서 데이터 분석 시 날짜 서식 오차 및 컬럼 인식 문제를 우회하여 안전하게 데이터를 추출합니다.")
    
    m_col1, m_col2 = st.columns([1, 2])
    
    with m_col1:
        st.markdown("### 📥 1. 마감 기본자료 업로드")
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
        st.markdown("### 📋 2. 외상매입금 대장 산출 프리뷰")
        if calc_ap_btn:
            if not uploaded_payable_plan:
                st.error("❌ 정산 처리를 위해 반입계획서 엑셀 파일을 업로드해 주세요.")
            else:
                with st.spinner(f"🤖 고유 키워드 추적 및 {selected_month_num}월 단일 대계 집계 테이블 구성 중..."):
                    try:
                        excel_file = pd.ExcelFile(uploaded_payable_plan)
                        # 시트명 대소문자 및 공백 제거 처리
                        sheet_names_lower = {s.lower().replace(" ", ""): s for s in excel_file.sheet_names}
                        
                        processed_list = []
                        
                        # ----------------------------------------------------
                        # ① NDP 시트 처리 (J, K, L, M열)
                        # ----------------------------------------------------
                        ndp_sheet_key = next((s for s in sheet_names_lower if "ndp" in s), None)
                        if ndp_sheet_key:
                            df_ndp = pd.read_excel(excel_file, sheet_name=sheet_names_lower[ndp_sheet_key], header=None)
                            start_parsing = False
                            c_lot, c_pname, c_pickup = 1, 2, 6  # 기본 인덱스
                            
                            for idx, row in df_ndp.iterrows():
                                # 헤더행 판단 시 동적 매칭 폭 확대
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
                                        
                                        # 고정 열 적용 (J=9, K=10, L=11, M=12)
                                        rl_val = float(str(row[9]).replace(",", "")) if pd.notna(row[9]) else 0.0
                                        kg_val = float(str(row[10]).replace(",", "")) if pd.notna(row[10]) else 0.0
                                        sqm_val = float(str(row[11]).replace(",", "")) if pd.notna(row[11]) else 0.0
                                        amt_val = float(str(row[12]).replace(",", "")) if pd.notna(row[12]) else 0.0
                                        
                                        fx_rate = get_hana_first_exrate(accounting_date)
                                        processed_list.append({
                                            "품명": p_name, "LOT No.": lot_val, "회계일자": accounting_date,
                                            "R/L": rl_val, "중량": kg_val, "면적": sqm_val, "Amount($)": amt_val,
                                            "환율": fx_rate, "원화금액": int(amt_val * fx_rate)
                                        })
                        
                        # ----------------------------------------------------
                        # ② ENSO 시트 처리 (H, I, J, K열)
                        # ----------------------------------------------------
                        enso_sheet_key = next((s for s in sheet_names_lower if "enso" in s), None)
                        if enso_sheet_key:
                            df_enso = pd.read_excel(excel_file, sheet_name=sheet_names_lower[enso_sheet_key], header=None)
                            start_parsing = False
                            c_lot, c_pname, c_arr = 1, 2, 4  # 기본 인덱스
                            
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
                                        
                                        # 고정 열 적용 (H=7, I=8, J=9, K=10)
                                        rl_val = float(str(row[7]).replace(",", "")) if pd.notna(row[7]) else 0.0
                                        kg_val = float(str(row[8]).replace(",", "")) if pd.notna(row[8]) else 0.0
                                        sqm_val = float(str(row[9]).replace(",", "")) if pd.notna(row[9]) else 0.0
                                        amt_val = float(str(row[10]).replace(",", "")) if pd.notna(row[10]) else 0.0
                                        
                                        fx_rate = get_hana_first_exrate(accounting_date)
                                        processed_list.append({
                                            "품명": p_name, "LOT No.": lot_val, "회계일자": accounting_date,
                                            "R/L": rl_val, "중량": kg_val, "면적": sqm_val, "Amount($)": amt_val,
                                            "환율": fx_rate, "원화금액": int(amt_val * fx_rate)
                                        })
                        
                        # ----------------------------------------------------
                        # ③ 통합 서식 빌드 및 총합계 적용
                        # ----------------------------------------------------
                        if not processed_list:
                            st.warning(f"⚠️ 업로드하신 반입계획서 내에서 선택하신 {selected_month_num}월 조건에 매칭되는 유효한 날짜가 인식되지 않았습니다. 파일 내부 날짜 컬럼을 확인해 주세요.")
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
                                
                                # 최하단 총합계 작성
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
                                
                            st.success(f"🎯 패치 적용 완료! {target_month} 외상매입금 대장이 정석 포맷에 맞춰 일괄 취합되었습니다.")
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
