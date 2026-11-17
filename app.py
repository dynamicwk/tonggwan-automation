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
st.set_page_config(
    layout="wide", 
    page_title="삼륭물산 구매무역팀 마감 포털"
)

def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_filename = "삼륭물산한글로고.png"
possible_paths = [
    logo_filename,
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
    page_bg_img = f"""
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
    """
    st.markdown(page_bg_img, unsafe_allow_html=True)

header_col1, header_col2 = st.columns([2, 1])
with header_col1:
    st.markdown(
        """
        <div style="padding-top: 10px;">
            <span style="font-size: 28px; font-weight: bold; color: #1e1e1e;">
                📊 삼륭물산 구매무역팀 통합 마감 포털
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )
with header_col2:
    st.markdown(
        """
        <div style="text-align: right; padding-top: 15px;">
            <span style="font-size: 15px; color: #333; font-weight: bold; 
                         background-color: rgba(255,255,255,0.9); 
                         padding: 6px 14px; border-radius: 6px; 
                         border: 1px solid #ddd;">
                🏢 삼륭물산 구매무역팀
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.caption("통관 월별납부 부가세 정산 | 해상물류비 마감감사 | 외상매입금 현황 마스터")
st.markdown("---")

# ==========================================
# 2. 공통 엔진 및 유틸리티 함수
# ==========================================
GEMINI_API_KEY = "AQ.Ab8RN6Le_B-K4XsTTGDe6Ny00O4JgZnb2uv2_xCKxpw6X0a_VQ"

@st.cache_data(ttl=86400)
def get_realtime_exchange_rate(date_str):
    try:
        t_date = pd.to_datetime(date_str).strftime("%Y%m%d")
        url_base = "https://www.kebhana.com/cms/rate/wpHanaIndex.do"
        url = f"{url_base}?searchDate={t_date}&searchInqCount=1"
        req = urllib.request.Request(
            url, 
            headers={'User-Agent': 'Mozilla/5.0'}
        )
        html = urllib.request.urlopen(req).read()
        soup = BeautifulSoup(html, 'html.parser')
        
        tables = soup.find_all('table')
        if tables:
            for row in tables[0].find_all('tr'):
                cells = row.find_all('td')
                if cells and 'USD' in str(cells[0]):
                    rate_txt = cells[8].text.strip().replace(",", "")
                    return float(rate_txt)
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
    if "200" in clean_lot:
        return "200ml"
    if "340" in clean_lot:
        return "340ml"
    if "500" in clean_lot:
        return "500ml"
    if "1000" in clean_lot:
        return "1000ml"
    if "HP" in clean_lot:
        return "HP"
    return str(fallback_val).strip()

def make_unique_cols(columns):
    cols = []
    count = {}
    for c in columns:
        c_str = str(c).strip()
        if c_str in count:
            count[c_
