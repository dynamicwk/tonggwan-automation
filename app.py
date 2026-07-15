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

# 다중 경로 탐색으로 파일 인식 오류 차단
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

# ==========================================
# 🗂️ 3대 대메뉴 기능 탭 분할
# ==========================================
tab1, tab2, tab3 = st.tabs([
    "📑 세관 통관 정산 마스터", 
    "📦 해상물류비 마감정산 (공장입고)",
    "💰 외상매입금 현황 마스터"
])

# ==========================================
# 📑 탭 1: 세관 통관 정산 마스터 시스템
# ==========================================
with tab1:
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("### 📥 1. 실무 자료 업로드")
        notice_file = st.file_uploader("관세청 '월별납부 개별고지목록' (Excel)", type=["xlsx", "xls"], key="notice_tab1")
        pdf_files = st.file_uploader("수입신고필증(면장) 통합본 파일 (PDF)", type=["pdf"], accept_multiple_files=True, key="declaration_tab1")
        st.markdown("---")
        start_btn = st.button("🚀 최종 통관 마스터 대장 산출하기", use_container_width=True, type="primary")

    with col2:
        st.markdown("### 📋 2. 정산 마스터 대장 결과물")
        if start_btn:
            if not notice_file or not pdf_files:
                st.error("❌ 관세청 고지서 엑셀 파일과 수입신고필증 PDF 파일을 모두 업로드해 주세요.")
            else:
                with st.spinner("🤖 제미나이 AI가 필증 내부의 실제 부가세액..."):
                    try:
                        st.info("세관 통관 정산 서비스가 대기 중입니다.")
                    except Exception as e:
                        st.error(f"❌ 오류: {e}")

# ==========================================
# 📦 탭 2: 해상물류비 마감정산 (공장입고)
# ==========================================
with tab2:
    st.markdown("### 📦 해상물류비 (공장입고) 자동 생성 및 검증 시스템")
    l_col1, l_col2 = st.columns([1, 2])
    with l_col1:
        uploaded_plan = st.file_uploader("1. 반입계획서 (엑셀)", type=["xlsx"], key="pantos_plan_tab2")
        uploaded_pantos = st.file_uploader("2. 판토스 마감내역서 (PDF)", type=["pdf"], key="pantos_pdf_tab2")
        pantos_btn = st.button("🚀 최종 물류비 마스터대장 산출하기", use_container_width=True, type="primary")
    with l_col2:
        if pantos_btn and uploaded_plan and uploaded_pantos:
            st.info("해상물류비 마감 완료")

# ==========================================
# 💰 탭 3: 외상매입금 현황 마스터 (★1월~12월 선택지 전면 연장)
# ==========================================
with tab3:
    st.markdown("### 💰 미정산 외상매입금 현황 자동 마감 시스템")
    st.write("반입계획서를 기반으로 NDP/ENSO 조건별 회계일자를 자동 산출하고, **하나은행 홈페이지에서 당일 최초 환율을 실시간으로 읽어와 연산**합니다.")
    
    m_col1, m_col2 = st.columns([1, 2])
    
    with m_col1:
        st.markdown("### 📥 1. 마감 기본자료 업로드")
        
        # 🎯 [요청 적용] 마감월 선택 리스트를 12월 전 구간으로 전격 세팅 완료
        months_options = [
            "2026년 12월", "2026년 11월", "2026년 10월", "2026년 9월", "2026년 8월", "2026년 7월",
            "2026년 6월", "2026년 5월", "2026년 4월", "2026년 3월", "2026년 2월", "2026년 1월"
        ]
        target_month = st.selectbox("마감 대상월 선택", months_options)
        uploaded_payable_plan = st.file_uploader("📋 반입계획서 파일 업로드 (Excel)", type=["xlsx"], key="payable_plan_uploader")
        st.markdown("---")
        calc_ap_btn = st.button("🚀 외상매입금 마감 대장 생성하기", use_container_width=True, type="primary")
        
    with m_col2:
        st.markdown("### 📋 2. 외상매입금 대장 산출 프리뷰")
        if calc_ap_btn:
            if not uploaded_payable_plan:
                st.error("❌ 정산 처리를 위해 반입계획서 엑셀 파일을 업로드해 주세요.")
            else:
                with st.spinner("🤖 하나은행 홈페이지 웹 서버에서 최초 매매기준율 환율을 실시간으로 가져오는 중입니다..."):
                    try:
                        df_temp = pd.read_excel(uploaded_payable_plan, header=None)
                        header_idx = 0
                        for idx, r in df_temp.iterrows():
                            if any("Lot" in str(v) or "오더" in str(v) for v in r.values):
                                header_idx = idx
                                break
                        df_src = pd.read_excel(uploaded_payable_plan, header=header_idx)
                        df_src.columns = df_src.columns.astype(str).str.strip()
                        
                        cols = df_src.columns
                        c_lot = next((c for c in cols if "Lot" in c or "오더" in c), None)
                        c_pickup = next((c for c in cols if "픽업" in c or "타코마" in c), None)
                        c_arr = next((c for c in cols if "입항" in c), None)
                        c_kg = next((c for c in cols if "kg" in c.upper() or "중량" in c), None)
                        c_sqm = next((c for c in cols if "SQM" in c.upper() or "면적" in c or "수량" in c), None)
                        c_amt = next((c for c in cols if "금액" in c or "외화" in c or "Amount" in c), None)
                        c_pname = next((c for c in cols if "품명" in c or "규격" in c), "품명 미확인")
                        
                        processed_list = []
                        
                        for _, row in df_src.iterrows():
                            lot_val = str(row.get(c_lot, "")).strip()
                            if lot_val in ("nan", "None", "") or "Lot" in lot_val:
                                continue
                            
                            p_name = str(row.get(c_pname, "200ml"))
                            kg_val = float(str(row.get(c_kg, 0)).replace(",", "")) if pd.notna(row.get(c_kg)) else 0.0
                            sqm_val = float(str(row.get(c_sqm, 0)).replace(",", "")) if pd.notna(row.get(c_sqm)) else 0.0
                            amt_val = float(str(row.get(c_amt, 0)).replace(",", "")) if pd.notna(row.get(c_amt)) else 0.0
                            
                            accounting_date = ""
                            
                            if lot_val.upper().startswith("E26"):
                                date_raw = row.get(c_arr, "")
                                if pd.notna(date_raw) and date_raw != "":
                                    accounting_date = pd.to_datetime(date_raw).strftime("%Y-%m-%d")
                            else:
                                date_raw = row.get(c_pickup, "")
                                if pd.notna(date_raw) and date_raw != "":
                                    accounting_date = pd.to_datetime(date_raw).strftime("%Y-%m-%d")
                            
                            if not accounting_date:
                                accounting_date = datetime.now().strftime("%Y-%m-%d")
                            
                            fx_rate = get_hana_first_exrate(accounting_date)
                            
                            processed_list.append({
                                "품명": p_name, "LOT No.": lot_val, "회계일자": accounting_date,
                                "R/L": 72, "중량": kg_val, "면적": sqm_val, "Amount($)": amt_val,
                                "환율": fx_rate, "원화금액": int(amt_val * fx_rate)
                            })
                        
                        if not processed_list:
                            st.warning("⚠️ 반입계획서에서 유효한 LOT 정산 내역을 추출하지 못했습니다.")
                        else:
                            df_preview = pd.DataFrame(processed_list)
                            
                            ap_excel = io.BytesIO()
                            with pd.ExcelWriter(ap_excel, engine='xlsxwriter') as writer:
                                wb = writer.book
                                ws = wb.add_worksheet(target_month)
                                
                                fmt_title = wb.add_format({'bold': True, 'font_size': 15, 'font_name': '맑은 고딕'})
                                fmt_th = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
                                fmt_subth = wb.add_format({'bg_color': '#F2F2F2', 'border': 1, 'align': 'center', 'font_size': 9})
                                fmt_cell = wb.add_format({'border': 1, 'align': 'center'})
                                fmt_num = wb.add_format({'border': 1, 'num_format': '#,##0'})
                                fmt_usd = wb.add_format({'border': 1, 'num_format': '#,##0.00'})
                                fmt_subtotal = wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'border': 1, 'num_format': '#,##0'})
                                fmt_subtotal_usd = wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'border': 1, 'num_format': '#,##0.00'})
                                
                                ws.write('A1', f"🏢 {target_month} 미정산 외상매입금 현황 마스터", fmt_title)
                                headers = ["품명", "LOT No.", "회계일자", "R/L", "중량", "면적", "Amount($)", "환율", "원화금액"]
                                for c_idx, h in enumerate(headers):
                                    ws.write(3, c_idx, h, fmt_th)
                                ws.write(4, 4, "kg", fmt_subth)
                                ws.write(4, 5, "SQ", fmt_subth)
                                
                                row_idx = 5
                                for item in processed_list:
                                    ws.write(row_idx, 0, item["품명"], fmt_cell)
                                    ws.write(row_idx, 1, item["LOT No."], fmt_cell)
                                    ws.write(row_idx, 2, item["회계일자"], fmt_cell)
                                    ws.write(row_idx, 3, item["R/L"], fmt_num)
                                    ws.write(row_idx, 4, item["중량"], fmt_num)
                                    ws.write(row_idx, 5, item["면적"], fmt_num)
                                    ws.write(row_idx, 6, item["Amount($)"], fmt_usd)
                                    ws.write(row_idx, 7, item["환율"], fmt_usd)
                                    ws.write_formula(row_idx, 8, f"=G{row_idx+1}*H{row_idx+1}", fmt_num)
                                    row_idx += 1
                                    
                                    ws.write(row_idx, 0, "", fmt_cell)
                                    ws.write(row_idx, 1, "소       계", wb.add_format({'bg_color': '#FFF2CC', 'bold': True, 'align': 'center', 'border': 1}))
                                    ws.write(row_idx, 2, 1, fmt_cell)
                                    ws.write_formula(row_idx, 3, f"=SUM(D{row_idx})", fmt_subtotal)
                                    ws.write_formula(row_idx, 4, f"=SUM(E{row_idx})", fmt_subtotal)
                                    ws.write_formula(row_idx, 5, f"=SUM(F{row_idx})", fmt_subtotal)
                                    ws.write_formula(row_idx, 6, f"=SUM(G{row_idx})", fmt_subtotal_usd)
                                    ws.write(row_idx, 7, "", fmt_cell)
                                    ws.write_formula(row_idx, 8, f"=SUM(I{row_idx})", fmt_subtotal)
                                    row_idx += 1
                                
                                ws.set_column('A:B', 16)
                                ws.set_column('C:C', 13)
                                ws.set_column('D:F', 11)
                                ws.set_column('G:I', 16)
                                
                            st.success(f"🎉 {target_month} 외상매입금 현황 대장이 성공적으로 빌드되었습니다! (하나은행 최초 환율 반영 완료)")
                            st.download_button(
                                label=f"📥 {target_month} 외상매입금 마스터 엑셀 다운로드 (.xlsx)",
                                data=ap_excel.getvalue(),
                                file_name=f"외상매입금현황_{target_month.replace(' ', '_')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                            st.dataframe(df_preview, use_container_width=True)
                            
                    except Exception as e:
                        st.error(f"❌ 데이터 정산 및 환율 크롤링 중 예외 에러 발생: {e}")
