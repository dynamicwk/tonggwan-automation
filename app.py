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

# 웹사이트 설정 및 디자인
st.set_page_config(layout="wide", page_title="삼륭물산 구매무역팀 마감 포털")

# 🖼️ [워터마크 엔진]
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_filename = "삼륭물산한글로고.png"
possible_paths = [logo_filename, os.path.join(os.path.dirname(__file__), logo_filename), os.path.join(os.getcwd(), logo_filename)]
bin_str = ""
for p in possible_paths:
    if os.path.exists(p):
        try:
            with open(p, "rb") as f:
                bin_str = base64.b64encode(f.read()).decode()
            break
        except Exception: pass

if bin_str:
    st.markdown(f'''<style>.stApp {{background-image: url("data:image/png;base64,{bin_str}"); background-size: 35%; background-repeat: no-repeat; background-position: center center; background-attachment: fixed;}} .block-container {{background-color: rgba(255, 255, 255, 0.88); border-radius: 12px; padding: 30px !important;}}</style>''', unsafe_allow_html=True)

# 🔑 [실시간 환율 수집 엔진] - 회계일자별 매매기준율 조회
@st.cache_data(ttl=86400) # 하루 동안 캐싱하여 중복 조회 방지
def get_realtime_exchange_rate(date_str):
    try:
        # 날짜 포맷팅 (YYYYMMDD)
        target_date = pd.to_datetime(date_str).strftime("%Y%m%d")
        url = f"https://www.kebhana.com/cms/rate/wpHanaIndex.do?searchDate={target_date}&searchInqCount=1"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'html.parser')
        
        # 하나은행 웹페이지 테이블에서 USD 매매기준율 추출
        rows = soup.select("table > tbody > tr")
        for row in rows:
            cols = row.find_all("td")
            if len(cols) > 8 and "USD" in cols[0].text:
                # 9번째 열(인덱스 8)이 매매기준율
                rate = float(cols[8].text.replace(",", ""))
                return rate
        return 1325.50 # 실패 시 디폴트값
    except:
        return 1325.50

# 비정형 한글 날짜 처리
def parse_flexible_month(date_val, target_month_num):
    if pd.isna(date_val) or str(date_val).strip() == "": return None
    try:
        dt = pd.to_datetime(date_val)
        return dt if dt.month == target_month_num else None
    except:
        str_val = str(date_val).strip()
        match = re.search(r"(\d+)월", str_val)
        if match and int(match.group(1)) == target_month_num:
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

# ==========================================
# 💰 탭 3: 외상매입금 현황 마스터 (실시간 환율 엔진 탑재)
# ==========================================
st.tabs(["📑 세관 통관", "📦 해상물류비", "💰 외상매입금"])
# (이후 생략... 탭 3 구현부)

# [탭 3 코드 내부의 환율 적용 로직 수정]
# 기존: fx_rate = get_hana_first_exrate(accounting_date)
# 변경: fx_rate = get_realtime_exchange_rate(accounting_date)
