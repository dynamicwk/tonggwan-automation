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

# 🖼️ [워터마크 엔진] 깃허브 저장소에 업로드된 PNG 파일을 읽어서 CSS 배경으로 주입
def get_base64_of_bin_file(bin_file):
    with open(bin_file, 'rb') as f:
        data = f.read()
    return base64.b64encode(data).decode()

logo_filename = "삼륭물산한글로고.png"

if os.path.exists(logo_filename):
    bin_str = get_base64_of_bin_file(logo_filename)
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
        background-color: rgba(255, 255, 255, 0.6);
        border-radius: 10px;
        padding: 30px !important;
    }}
    </style>
    '''
    st.markdown(page_bg_img, unsafe_allow_html=True)

# 🏢 [상단 헤더 레이아웃]
header_col1, header_col2 = st.columns([2, 1])

with header_col1:
    st.markdown(
        """
        <div style="font-family: 'Malgun Gothic', sans-serif; padding-top: 10px;">
            <span style="font-size: 26px; font-weight: bold; color: #1e1e1e;">📊 삼륭물산 통관 정산 시스템</span>
        </div>
        """, 
        unsafe_allow_html=True
    )

with header_col2:
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
        key="notice_final_fixed_v18"
    )
    
    pdf_files = st.file_uploader(
        "수입신고필증(면장) 통합본 파일 (PDF)", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key="declaration_final_fixed_v18"
    )
    
    st.markdown("---")
    start_btn = st.button("🚀 최종 마스터 대장 산출하기", use_container_width=True, type="primary")
    
    if not os.path.exists(logo_filename):
        st.info("💡 '삼륭물산한글로고.png' 파일을 app.py와 같은 폴더에 업로드하시면 배경 워터마크 로고가 활성화됩니다.")

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
                    
                    # 🛠️ [프롬프트 극강 보강] ⑤⑦ 운임, ⑤⑧ 보험료 타겟팅 강화 조항 신설
                    prompt = """
                    당신은 관세 법인 소속의 정산 자동화 AI입니다. 제공된 수입신고필증 PDF 문서 전체를 전수조사하여, 각 '신고번호'별로 아래 항목들을 정확하게 추출하여 JSON 배열 형태로 응답해 주세요.
                    
                    [핵심 요구사항 - 반드시 준수]
                    1. usd_amount (결제금액): 소수점 이하 자리(센트 단위)가 있다면 절대로 버림/반올림하지 말고 원래의 소수점 둘째 자리까지 완벽한 소수로 추출하십시오. (예: 139504.48)
                    2. freight (⑤⑦ 운임): 각 란의 쪼개진 운임이 아니라, 필증 맨 아래 또는 결제금액 우측에 명시된 해당 신고번호 건의 '총 운임비 KRW 금액 전체'를 추출하십시오. 
                    3. insurance (⑤⑧ 보험료): 각 란의 소액 보험료가 아니라, 필증 맨 아래쪽에 명시된 해당 신고건의 '총 보험증권 가액/보험료 KRW 금액 전체'를 추출하십시오. 만약 인도조건이 CIF, DAP, CIP인 경우 특약이 없다면 기본적으로 0으로 처리하되, 필증에 총액 수치가 명시되어 있다면 그 수치를 우선합니다.
                    
                    추출 항목 리스트:
                    - shin_no: 신고번호
                    - shin_date: 신고일자
                    - bl_no: ④ B/L(AWB)번호
                    - fx_rate: 환율
                    - incoterms: 인도조건 (FCA, CIF, DAP, CIP 등)
                    - usd_amount: 결제금액 (소수점 포함)
                    - freight: 총 운임비 (KRW 금액 숫자만)
                    - insurance: 총 보험료 (KRW 금액 숫자만)
                    
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
                            
                            # 데이터 안전 정제
                            try:
                                freight_num = int(float(str(ai_data.get('freight', 0)).replace(",", "")))
                            except:
                                freight_num = 0
                                
                            try:
                                insurance_num = int(float(str(ai_data.get('insurance', 0)).replace(",", "")))
                            except:
                                insurance_num = 0
                                
                            if incoterms_type in ["CIF", "DAP", "CIP"] and freight_num == 0 and insurance_num == 0:
                                pass # 이미 0인 경우 유지
                                
                            note = "정상 매칭"
                        
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
                            "비고": note
                        })
                    
                    output_excel = io.BytesIO()
                    workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                    
                    df_final = pd.DataFrame(processed_data)
                    df_final.to_excel(workbook, sheet_name="통관월납_정산대장", index=False)
                    
                    ws = workbook.sheets["통관월납_정산대장"]
                    wb = workbook.book
                    
                    num_format = wb.add_format({'num_format': '#,##0'})
                    usd_format = wb.add_format({'num_format': '#,##0.00'})
                    fx_format = wb.add_format({'num_format': '#,##0.0000'})
                    align_center = wb.add_format({'align': 'center'})
                    
                    ws.set_column('D:D', 18, num_format)
                    ws.set_column('G:G', 18, usd_format)
                    ws.set_column('I:I', 12, fx_format)
                    ws.set_column('J:K', 15, num_format)
                    
                    ws.write('N1', '과세가격(원화산출식)')
                    ws.write('O1', '수식검증 부가세(원단위 버림)')
                    ws.write('P1', '고지액 검증 결과')
                    
                    for i in range(2, len(processed_data) + 2):
                        ws.write_formula(f'N{i}', f'=(G{i}*I{i})+J{i}+K{i}', num_format)
                        ws.write_formula(f'O{i}', f'=ROUNDDOWN(N{i}*0.1, -1)', num_format)
                        ws.write_formula(f'P{i}', f'=IF(D{i}=O{i}, "일치", "❌ 금액 불일치")', align_center)
                        
                    workbook.close()
                    excel_data = output_excel.getvalue()
                    
                    st.success("🎉 정산 마스터 대장 작성이 완료되었습니다!")
                    
                    st.download_button(
                        label="📥 정산 마스터 대장 엑셀 파일 다운로드 (.xlsx)",
                        data=excel_data,
                        file_name="통관월납_정산_마스터대장_수식포함.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.dataframe(df_final, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ 최종 양식 변환 중 오류가 발생했습니다: {e}")
