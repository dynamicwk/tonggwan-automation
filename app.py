import streamlit as st
import pandas as pd
import math
from google import genai
from google.genai import types
import io
import json
import os
import base64

# 웹사이트 설정 및 디자인 (넓은 화면 모드)
st.set_page_config(layout="wide")

# 🖼️ [워터마크 엔진] 저장소의 PNG 파일을 읽어서 CSS 배경으로 주입하는 함수
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_filename = "삼륭물산한글로고.png"

if os.path.exists(logo_filename):
    bin_str = get_base64_of_bin_file(logo_filename)
    # 투명도(opacity) 0.06 정도로 아주 은은하게 설정하여 본문 글씨를 가리지 않도록 중앙 배치
    page_bg_img = f'''
    <style>
    .stApp {{
        background-image: url("data:image/png;base64,{bin_str}");
        background-size: 55%; /* 화면 대비 로고 크기 설정 (큼직하게 55%) */
        background-repeat: no-repeat;
        background-position: center center;
        background-attachment: fixed;
        opacity: 1.0; /* 본문 투명도 고정 */
    }}
    /* 콘텐츠 박스가 배경 로고를 가리지 않도록 메인 영역 배경 투명화 처리 */
    .block-container {{
        background-color: rgba(255, 255, 255, 0.4);
        border-radius: 10px;
        padding: 30px !important;
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

# 🏢 [상단 레이아웃] 좌측 상단 타이틀 및 우측 상단 소속 표시
header_col1, header_col2 = st.columns([2, 1])

with header_col1:
    st.markdown(
        """
        <div style="font-family: 'Malgun Gothic', sans-serif; padding-top: 10px; display: flex; align-items: center; gap: 15px;">
            <span style="font-size: 26px; font-weight: bold; color: #1e1e1e;">📊 삼륭물산 통관 정산 시스템</span>
        </div>
        """, 
        unsafe_allow_html=True
    )

with header_col2:
    # 우측 상단 삼륭물산 구매무역팀 깔끔한 텍스트 처리
    st.markdown(
        """
        <div style="text-align: right; font-family: 'Malgun Gothic', sans-serif; padding-top: 15px;">
            <span style="font-size: 15px; color: #333333; font-weight: bold;">
                🏢 삼륭물산 구매무역팀
            </span>
        </div>
        """, 
        unsafe_allow_html=True
    )

st.caption("안산 / 안양 / 부산 전 세관 고지서 원본 취합 및 수입신고필증 PDF 실시간 데이터 교차 검증 엔진")
st.markdown("<hr style='margin-top: 5px; margin-bottom: 25px; border-top: 1px solid #eee;'>", unsafe_allow_html=True)

# 🔑 구글 제미나이 API 키 고정
GEMINI_API_KEY = "AQ.Ab8RN6Le_B-K4XsTTGDe6Ny00O4JgZnb2uv2_xCKxpw6X0a_VQ"

# 좌측 업로드 공간, 우측 결과창 분할
col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### 📥 1. 실무 자료 업로드")
    
    notice_file = st.file_uploader(
        "관세청 '월별납부 개별고지목록' (Excel)", 
        type=["xlsx", "xls"], 
        key="notice_final_fixed_v13"
    )
    
    pdf_files = st.file_uploader(
        "수입신고필증(면장) 통합본 파일 (PDF)", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key="declaration_final_fixed_v13"
    )
    
    st.markdown("---")
    start_btn = st.button("🚀 최종 마스터 대장 산출하기", use_container_width=True, type="primary")
    
    if not os.path.exists(logo_filename):
        st.info("💡 '삼륭물산한글로고.png' 파일을 app.py와 같은 폴더에 업로드하시면 중앙 배경에 은은한 워터마크 로고가 활성화됩니다.")

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
                    
                    - shin_no: 신고번호
                    - shin_date: 신고일자
                    - bl_no: ④ B/L(AWB)번호
                    - fx_rate: 환율
                    - incoterms: 인도조건 (FCA, CIF, DAP 등)
                    - usd_amount: 결제금액
                    - freight: 운임
                    - insurance: 보험료
                    
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
                    pdf_master_dict = { "".join(filter(str.isalnum, item['shin_no'])): item for item in extracted_data }
                    
                    processed_data = []
                    
                    for idx, row in df_notice_all.iterrows():
