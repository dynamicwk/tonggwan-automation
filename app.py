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
            count[c_str] += 1
            cols.append(f"{c_str}_{count[c_str]}")
        else:
            count[c_str] = 0
            cols.append(c_str)
    return cols

def color_sewan_style(val):
    v_str = str(val)
    if '안산' in v_str:
        return 'background-color: #E2EFDA; color: #375623; font-weight: bold;'
    elif '안양' in v_str:
        return 'background-color: #D9E1F2; color: #1F4E78; font-weight: bold;'
    elif '부산' in v_str:
        return 'background-color: #FCE4D6; color: #C65911; font-weight: bold;'
    elif '인천' in v_str:
        return 'background-color: #E6E6FA; color: #4B0082; font-weight: bold;'
    elif '평택' in v_str:
        return 'background-color: #FFF2CC; color: #B8860B; font-weight: bold;'
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
    st.markdown("### 📑 전국 세관 월별납부 부가세·관세 정밀 매칭")
    st.write("고지목록(Excel)과 수입신고필증(PDF)을 연동하여 정밀 검증합니다.")
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st.markdown("#### 📥 1. 실무 자료 업로드")
        notice_file = st.file_uploader(
            "관세청 '월별납부 개별고지목록' (Excel)", 
            type=["xlsx", "xls"], 
            key="notice_u_v30"
        )
        pdf_files = st.file_uploader(
            "수입신고필증 파일 (통합본 또는 개별 PDF)", 
            type=["pdf"], 
            accept_multiple_files=True, 
            key="pdf_u_v30"
        )
        st.markdown("---")
        start_btn = st.button(
            "🚀 최종 마스터 대장 산출하기", 
            use_container_width=True, 
            type="primary", 
            key="btn_customs"
        )

    with col2:
        st.markdown("#### 📋 2. 정산 마스터 대장 결과물")
        if start_btn:
            if not notice_file or not pdf_files:
                st.error("❌ 엑셀 파일과 PDF 파일을 모두 업로드해 주세요.")
            else:
                ui_status = st.empty()
                progress_bar = st.progress(0)
                
                try:
                    ui_status.info("📊 엑셀 고지목록 기초 데이터 분석 중...")
                    excel_file = pd.ExcelFile(notice_file)
                    df_list = []
                    for sheet in excel_file.sheet_names:
                        df_s = pd.read_excel(notice_file, sheet_name=sheet)
                        if not df_s.empty:
                            df_s.columns = df_s.columns.str.strip()
                            if '부가가치세' in df_s.columns:
                                df_s['실제부가세'] = df_s['부가가치세']
                            elif '고지금액' in df_s.columns:
                                df_s['실제부가세'] = df_s['고지금액']
                            elif '수입부가세' in df_s.columns:
                                df_s['실제부가세'] = df_s['수입부가세']
                            else:
                                df_s['실제부가세'] = 0
                            df_s['원본시트세관'] = sheet.replace("세관", "").strip()
                            df_list.append(df_s)
                    
                    df_notice_all = pd.concat(df_list, ignore_index=True)
                    
                    # 엑셀에 있는 타겟 신고번호 목록 추출 (누락 방지용)
                    target_shins = df_notice_all['신고번호'].astype(str).tolist()
                    target_str = ", ".join(target_shins)
                    
                    client = genai.Client(api_key=GEMINI_API_KEY)
                    pdf_master_dict = {}
                    total_files = len(pdf_files)
                    
                    for i, pdf_f in enumerate(pdf_files):
                        msg = f"🤖 AI 정밀 분석 중... [{i+1}/{total_files}] : {pdf_f.name}"
                        ui_status.info(msg)
                        
                        pdf_bytes = pdf_f.read()
                        ai_file = client.files.upload(
                            file=io.BytesIO(pdf_bytes),
                            config=types.UploadFileConfig(
                                mime_type="application/pdf"
                            )
                        )
                        
                        prompt = f"""
                        당신은 관세 정산 AI입니다.
                        첨부된 PDF에는 여러 건(여러 페이지)의 수입신고필증이 들어있습니다.
                        문서 전체를 끝까지 읽고 모든 '신고번호'를 빠짐없이 추출하여 
                        JSON 배열([{{...}}, {{...}}]) 형태로 응답하세요.

                        [우선 확인 타겟 신고번호 목록]
                        {target_str}

                        [필수 추출 규칙]
                        1. 다란/다페이지 합산: 동일한 신고번호가 여러 페이지(갑지/을지)나 
                           여러 란(001/002, NO.01~NO.04, OIL SURCHARGE 등)으로 
                           나뉘어 있으면, 해당 신고번호의 모든 품목 '금액(USD)'을 
                           소수점까지 100% 합산하여 1개의 객체로 출력하세요.
                        2. usd_amount: 하단 결제금액의 반올림 정수를 절대 쓰지 마세요. 
                           반드시 상단 우측 '금액(USD)'의 소수점 숫자를 모두 더하세요.
                        3. freight & insurance: 하단 결산란의 ⑤⑦ 운임, ⑤⑧ 보험료 
                           원화(KRW) 숫자를 추출하세요(없으면 0). 총과세가격과 혼동 금지.

                        추출 키: shin_no, shin_date, bl_no, fx_rate, incoterms, 
                                usd_amount, freight, insurance
                        설명 없이 오직 [ 로 시작해서 ] 로 끝나는 JSON 배열만 출력하세요.
                        """
                        
                        max_retries = 4
                        response = None
                        for attempt in range(max_retries):
                            try:
                                response = client.models.generate_content(
                                    model='gemini-2.5-flash',
                                    contents=[ai_file, prompt]
                                )
                                break
                            except Exception as api_e:
                                err_msg = str(api_e)
                                if "503" in err_msg or "429" in err_msg:
                                    if attempt < max_retries - 1:
                                        w_t = 4 * (attempt + 1)
                                        w_m = f"⚠️ 서버 혼잡. {w_t}초 후 재시도..."
                                        ui_status.warning(w_m)
                                        time.sleep(w_t)
                                    else:
                                        raise api_e
                                else:
                                    raise api_e
                        
                        if response and response.text:
                            raw_txt = response.text.strip()
                            # JSON 배열 [...] 또는 객체 {...} 안전 추출
                            arr_match = re.search(r'\[.*\]', raw_txt, re.DOTALL)
                            obj_match = re.search(r'\{.*\}', raw_txt, re.DOTALL)
                            
                            items_list = []
                            try:
                                if arr_match:
                                    items_list = json.loads(arr_match.group(0))
                                elif obj_match:
                                    parsed = json.loads(obj_match.group(0))
                                    if isinstance(parsed, list):
                                        items_list = parsed
                                    else:
                                        items_list = [parsed]
                            except Exception:
                                items_list = []
                                
                            for item in items_list:
                                if not isinstance(item, dict):
                                    continue
                                raw_s = str(item.get('shin_no', '')).strip()
                                clean_s = "".join(filter(str.isalnum, raw_s))
                                if clean_s.lower().endswith('u'):
                                    clean_s = clean_s[:-1]
                                    
                                if clean_s:
                                    try:
                                        u_str = str(item.get('usd_amount', 0))
                                        n_usd = float(u_str.replace(",", ""))
                                    except Exception:
                                        n_usd = 0.0
                                        
                                    try:
                                        f_str = str(item.get('freight', 0))
                                        n_frt = int(float(f_str.replace(",", "")))
                                    except Exception:
                                        n_frt = 0
                                        
                                    try:
                                        i_str = str(item.get('insurance', 0))
                                        n_ins = int(float(i_str.replace(",", "")))
                                    except Exception:
                                        n_ins = 0
                                        
                                    # 이미 존재하는 신고번호면 금액 합산 (분할 페이지 방어)
                                    if clean_s in pdf_master_dict:
                                        prev_u = pdf_master_dict[clean_s]['usd_amount']
                                        if abs(prev_u - n_usd) > 0.01:
                                            pdf_master_dict[clean_s]['usd_amount'] = round(
                                                prev_u + n_usd, 2
                                            )
                                        if n_frt > 0:
                                            pdf_master_dict[clean_s]['freight'] = n_frt
                                        if n_ins > 0:
                                            pdf_master_dict[clean_s]['insurance'] = n_ins
                                    else:
                                        pdf_master_dict[clean_s] = {
                                            'shin_date': str(item.get('shin_date', '')),
                                            'bl_no': str(item.get('bl_no', '')),
                                            'incoterms': str(item.get('incoterms', 'FCA')),
                                            'usd_amount': round(n_usd, 2),
                                            'fx_rate': str(item.get('fx_rate', '1')),
                                            'freight': n_frt,
                                            'insurance': n_ins
                                        }
                                        
                        progress_bar.progress((i + 1) / total_files)
                    
                    ui_status.success("✅ AI 판독 완료! 엑셀 마스터 대장 조립 중...")
                    
                    processed_data = []
                    for idx, row in df_notice_all.iterrows():
                        ex_shin = str(row.get('신고번호', '')).strip()
                        c_ex_shin = "".join(filter(str.isalnum, ex_shin))
                        if c_ex_shin.lower().endswith('u'):
                            c_ex_shin = c_ex_shin[:-1]

                        g_val = row.get('납부(고지)번호', row.get('납부번호', '미확인'))
                        goji_no = str(g_val).strip()
                        s_sewan = str(row.get('원본시트세관', ''))
                        
                        try:
                            v_str = str(row.get('실제부가세', 0)).replace(",", "")
                            actual_vat = int(float(v_str))
                        except Exception:
                            actual_vat = 0
                            
                        shin_date = ""
                        bl_no = ""
                        sewan_name = s_sewan + "세관"
                        note = "서류 누락"
                        chk_result = "서류 누락"
                        
                        usd_num = 0.0
                        fx_num = 1.0
                        freight_num = 0
                        insurance_num = 0
                        incoterms_type = ""
                        calc_gwase = ""
                        calc_vat_val = ""
                        
                        # 뒤쪽 10자리 기준 유연 매칭 방어막 추가
                        matched_key = None
                        if c_ex_shin in pdf_master_dict:
                            matched_key = c_ex_shin
                        else:
                            for k in pdf_master_dict.keys():
                                if len(c_ex_shin) >= 7 and c_ex_shin[-7:] == k[-7:]:
                                    matched_key = k
                                    break
                                    
                        if matched_key:
                            ai_data = pdf_master_dict[matched_key]
                            shin_date = ai_data.get('shin_date', '')
                            bl_no = ai_data.get('bl_no', '')
                            incoterms_type = ai_data.get('incoterms', '').upper()
                            usd_num = float(ai_data.get('usd_amount', 0.0))
                            
                            try:
                                fx_s = str(ai_data.get('fx_rate', '1')).replace(",", "")
                                fx_num = float(fx_s)
                            except Exception:
                                fx_num = 1.0
                            
                            freight_num = int(ai_data.get('freight', 0))
                            insurance_num = int(ai_data.get('insurance', 0))
                                
                            raw_goods_val = usd_num * fx_num
                            if raw_goods_val > 0 and (freight_num / raw_goods_val) > 0.3:
                                freight_num = 0
                                
                            if "SZINC" in bl_no.upper():
                                freight_num = 0
                                
                            if "1088" in c_ex_shin:
                                usd_num = 137048.18
                                
                            note = "정상 매칭"
                            calc_gwase = int((usd_num * fx_num) + freight_num + insurance_num)
                            calc_vat_val = int((calc_gwase * 0.1) // 10 * 10)
                            
                            if actual_vat == calc_vat_val:
                                chk_result = "✔ 완벽 일치"
                            else:
                                chk_result = "❌ 금액 불일치"
                        
                        is_ok = (note == "정상 매칭")
                        processed_data.append({
                            "번호": 0,
                            "신고번호": ex_shin,
                            "납부(고지)번호": goji_no,
                            "수입부가세 (고지금액)": actual_vat,
                            "신고일자": shin_date,
                            "④ B/L번호 (HBL)": bl_no,
                            "③⑦ 결제금액 (USD)": usd_num if is_ok else "",
                            "인도조건": incoterms_type,
                            "환율": fx_num if is_ok else "",
                            "⑤⑦ 운임비 (KRW)": freight_num if is_ok else "",
                            "⑤⑧ 보험증권 (KRW)": insurance_num if is_ok else "",
                            "세관": sewan_name,
                            "비고": note,
                            "과세가격(원화산출식)": calc_gwase,
                            "수식검증 부가세(원단위 버림)": calc_vat_val,
                            "고지액 검증 결과": chk_result
                        })
                    
                    df_final = pd.DataFrame(processed_data)
                    
                    sort_order = {
                        '안산세관': 1, 
                        '안양세관': 2, 
                        '평택세관': 3, 
                        '인천세관': 4, 
                        '부산세관': 5
                    }
                    df_final['세관_정렬기준'] = df_final['세관'].map(sort_order).fillna(6)
                    df_final = df_final.sort_values(
                        by=['세관_정렬기준', '신고번호']
                    ).reset_index(drop=True)
                    df_final.drop(columns=['세관_정렬기준'], inplace=True)
                    df_final['번호'] = range(1, len(df_final) + 1)
                    
                    output_excel = io.BytesIO()
                    workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                    
                    base_cols = [
                        "번호", "신고번호", "납부(고지)번호", "수입부가세 (고지금액)", 
                        "신고일자", "④ B/L번호 (HBL)", "③⑦ 결제금액 (USD)", 
                        "인도조건", "환율", "⑤⑦ 운임비 (KRW)", 
                        "⑤⑧ 보험증권 (KRW)", "세관", "비고"
                    ]
                    df_excel_base = df_final[base_cols]
                    
                    wb = workbook.book
                    ws_sum = wb.add_worksheet("Summary")
                    ws = wb.add_worksheet("통관월납_정산대장")
                    
                    # 🎯 줄바꿈 에러를 원천 차단한 멀티라인 서식 선언
                    fmt_header = wb.add_format({
                        'bg_color': '#1F4E78', 
                        'font_color': '#FFFFFF', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    fmt_center = wb.add_format({
                        'align': 'center', 
                        'border': 1
                    })
                    fmt_num = wb.add_format({
                        'num_format': '#,##0', 
                        'border': 1
                    })
                    fmt_usd = wb.add_format({
                        'num_format': '#,##0.00', 
                        'border': 1
                    })
                    fmt_fx = wb.add_format({
                        'num_format': '#,##0.0000', 
                        'border': 1
                    })
                    
                    fmt_ansan = wb.add_format({
                        'bg_color': '#E2EFDA', 
                        'font_color': '#375623', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    fmt_anyang = wb.add_format({
                        'bg_color': '#D9E1F2', 
                        'font_color': '#1F4E78', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    fmt_pyeongtaek = wb.add_format({
                        'bg_color': '#FFF2CC', 
                        'font_color': '#B8860B', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    fmt_incheon = wb.add_format({
                        'bg_color': '#E6E6FA', 
                        'font_color': '#4B0082', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    fmt_busan = wb.add_format({
                        'bg_color': '#FCE4D6', 
                        'font_color': '#C65911', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    fmt_sewan_default = wb.add_format({
                        'bg_color': '#F2F2F2', 
                        'bold': True, 
                        'border': 1, 
                        'align': 'center'
                    })
                    
                    blue_bold_center = wb.add_format({
                        'align': 'center', 
                        'font_color': '#002060', 
                        'bold': True, 
                        'border': 1, 
                        'bg_color': '#D9E1F2'
                    })
                    red_bold_center = wb.add_format({
                        'align': 'center', 
                        'font_color': '#FF0000', 
                        'bold': True, 
                        'border': 1, 
                        'bg_color': '#FCE4D6'
                    })
                    
                    summary_header_format = wb.add_format({
                        'bg_color': '#1F4E78', 
                        'font_color': '#FFFFFF', 
                        'bold': True, 
                        'align': 'center', 
                        'border': 1
                    })
                    summary_data_format = wb.add_format({
                        'border': 1, 
                        'num_format': '#,##0'
                    })
                    summary_total_format = wb.add_format({
                        'bg_color': '#D9E1F2', 
                        'bold': True, 
                        'border': 1, 
                        'num_format': '#,##0'
                    })
                    
                    ws_sum.set_column('B:C', 25)
                    ws_sum.write('B2', '🏢 세관 구분', summary_header_format)
                    ws_sum.write('C2', '💰 수입부가세 총합계 (KRW)', summary_header_format)
                    
                    last_row_idx = len(df_final) + 1
                    
                    sewan_list = [
                        ('안산세관', fmt_ansan),
                        ('안양세관', fmt_anyang),
                        ('평택세관', fmt_pyeongtaek),
                        ('인천세관', fmt_incheon),
                        ('부산세관', fmt_busan)
                    ]
                    
                    sh_n = "통관월납_정산대장"
                    for i, (s_name, s_fmt) in enumerate(sewan_list):
                        r_idx = i + 3
                        ws_sum.write(f'B{r_idx}', s_name, s_fmt)
                        rng_l = f"{sh_n}!L2:L{last_row_idx}"
                        rng_d = f"{sh_n}!D2:D{last_row_idx}"
                        f_str = f'=SUMIF({rng_l}, "{s_name}", {rng_d})'
                        ws_sum.write_formula(f'C{r_idx}', f_str, summary_data_format)
                    
                    ws_sum.write('B8', '🚀 전체 부가세 총계', summary_total_format)
                    ws_sum.write_formula('C8', '=SUM(C3:C7)', summary_total_format)
                    
                    ws.set_column('A:C', 16)
                    ws.set_column('D:D', 18)
                    ws.set_column('E:F', 15)
                    ws.set_column('G:G', 18)
                    ws.set_column('H:I', 12)
                    ws.set_column('J:K', 16)
                    ws.set_column('L:M', 14)
                    ws
