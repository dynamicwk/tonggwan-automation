import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter

# 페이지 설정
st.set_page_config(layout="wide", page_title="삼륭물산 구매무역팀 마감 포털")

# [핵심 엔진 1] 실시간 환율 수집 (하나은행)
@st.cache_data(ttl=86400)
def get_realtime_exchange_rate(date_str):
    try:
        target_date = pd.to_datetime(date_str).strftime("%Y%m%d")
        url = f"https://www.kebhana.com/cms/rate/wpHanaIndex.do?searchDate={target_date}&searchInqCount=1"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'html.parser')
        rows = soup.select("table > tbody > tr")
        for row in rows:
            cols = row.find_all("td")
            if len(cols) > 8 and "USD" in cols[0].text:
                return float(cols[8].text.replace(",", ""))
        return 1325.50
    except: return 1325.50

# [핵심 엔진 2] 유틸리티
def parse_flexible_month(date_val, target_month_num):
    try:
        dt = pd.to_datetime(date_val)
        return dt if dt.month == target_month_num else None
    except:
        match = re.search(r"(\d+)월", str(date_val))
        if match and int(match.group(1)) == target_month_num: return datetime(2026, target_month_num, 1)
        return None

def auto_define_pname(lot_str, fallback_val):
    clean_lot = str(lot_str).strip().upper()
    if "200" in clean_lot: return "200ml"
    if "340" in clean_lot: return "340ml"
    if "500" in clean_lot: return "500ml"
    if "1000" in clean_lot: return "1000ml"
    if "HP" in clean_lot: return "HP"
    return str(fallback_val).strip()

# 탭 메뉴
tab1, tab2, tab3 = st.tabs(["📑 세관 통관 정산 마스터", "📦 해상물류비 마감정산", "💰 외상매입금 현황 마스터"])

with tab1:
    st.header("📑 세관 통관 정산 시스템")
    st.write("관세청 데이터와 수입신고필증을 매칭하여 정산합니다.")
    # (세관 정산 로직 구동 구간)

with tab2:
    st.header("📦 해상물류비 감사 시스템")
    st.write("판토스 마감내역과 반입계획서를 비교 검증합니다.")
    # (물류비 정산 로직 구동 구간)

with tab3:
    st.header("💰 외상매입금 현황 마스터")
    target_month = st.selectbox("마감 대상월", [f"2026년 {i}월" for i in range(12, 0, -1)])
    selected_month_num = int(re.search(r"(\d+)월", target_month).group(1))
    uploaded_plan = st.file_uploader("반입계획서 업로드", type=["xlsx"])
    
    if st.button("🚀 최종 마감 대장 생성"):
        if uploaded_plan:
            with st.spinner("실시간 환율 수집 및 마감 대장 생성 중..."):
                excel_file = pd.ExcelFile(uploaded_plan)
                processed_list = []
                # (중략: NDP/ENSO 데이터 파싱 로직)
                # 추출 시 fx_rate = get_realtime_exchange_rate(accounting_date) 적용
                st.success("실시간 환율이 적용된 마감 대장이 생성되었습니다.")
