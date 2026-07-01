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

# 관세청 파일과 면장 PDF 파일을 올리는 대시보드
col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### 📥 1. 실무 자료 업로드")
    
    notice_file = st.file_uploader(
        "관세청 '월별납부 개별고지목록' (Excel)", 
        type=["xlsx", "xls"], 
        key="notice_final_fixed_v21"
    )
    
    pdf_files = st.file_uploader(
        "수입신고필증(면장) 통합본 파일 (PDF)", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key="declaration_final_fixed_v21"
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
                    
                    prompt = """
                    당신은 관세 법인 소속의 정산 자동화 AI입니다. 제공된 수입신고필증 PDF 문서 전체를 전수조사하여, 각 '신고번호'별로 아래 항목들을 정확하게 추출하여 JSON 배열 형태로 응답해 주세요.
                    
                    [엄격한 추출 규칙 - 필수 준수]
                    1. 페이지 범위 격리: 여러 건의 면장이 합쳐진 문서이므로, 추출하려는 '신고번호'가 기재된 페이지 블록 내의 숫자만 참조하십시오. 다른 신고건의 페이지에 있는 숫자를 가져오면 절대 안 됩니다.
                    2. usd_amount (결제금액): 소수점 이하 자리(센트 단위)가 존재한다면 반올림하거나 정수화하지 마십시오. 소수점 둘째 자리까지의 원래의 값을 완벽한 소수(Float)로 추출하십시오. (예: 139504.48)
                    3. freight (⑤⑦ 운임) & insurance (⑤⑧ 보험료): 
                       - 해당 신고번호가 적힌 면장 양식 우측 하단 혹은 총액 기재 영역의 '총 금액'을 원화(KRW) 기준으로 추출하십시오.
                       - 만약 해당 신고건 면장의 총 운임비 또는 보험료 란에 아무런 금액이 적혀있지 않거나 공란(0 또는 비어있음)인 경우, 혹은 인도조건에 따라 수치가 존재하지 않는 샘플 건은 절대로 임의의 숫자를 유추하지 말고 반드시 0으로 명시하십시오.
                    
                    추출 항목 리스트:
                    - shin_no: 신고번호 (하이픈 포함 원래 형태 유지)
                    - shin_date: 신고일자 (YYYY/MM/DD)
                    - bl_no: ④ B/L(AWB)번호
                    - fx_rate: 환율
                    - incoterms: 인도조건 (FCA, CIF, DAP, CIP 등)
                    - usd_amount: 결제금액 (소수점 포함)
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
                            
                            try:
                                freight_num = int(float(str(ai_data.get('freight', 0)).replace(",", "")))
                            except:
                                freight_num = 0
                                
                            try:
                                insurance_num = int(float(str(ai_data.get('insurance', 0)).replace(",", "")))
                            except:
                                insurance_num = 0
                                
                            if "SZINC" in bl_no.upper():
                                freight_num = 0
                                
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
                    
                    # 🛠️ 수식 포함 다중 시트 엑셀 파일 작성 시작
                    output_excel = io.BytesIO()
                    workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                    
                    df_final = pd.DataFrame(processed_data)
                    
                    # 스타일 서식 정의
                    wb = workbook.book
                    num_format = wb.add_format({'num_format': '#,##0'})
                    usd_format = wb.add_format({'num_format': '#,##0.00'})
                    fx_format = wb.add_format({'num_format': '#,##0.0000'})
                    align_center = wb.add_format({'align': 'center'})
                    
                    # 🎨 [디자인 고도화] 일치 시 파란색 글씨 처리 스타일 추가
                    blue_bold_center = wb.add_format({'align': 'center', 'font_color': '#002060', 'bold': True})
                    red_bold_center = wb.add_format({'align': 'center', 'font_color': '#FF0000', 'bold': True})
                    
                    summary_header_format = wb.add_format({
                        'bg_color': '#1F4E78', 'font_color': '#FFFFFF', 'bold': True, 'align': 'center', 'border': 1
                    })
                    summary_data_format = wb.add_format({'border': 1, 'num_format': '#,##0'})
                    summary_total_format = wb.add_format({'bg_color': '#D9E1F2', 'bold': True, 'border': 1, 'num_format': '#,##0'})
                    
                    # --- [시트 1 : Summary 시트 생성] ---
                    ws_sum = wb.add_sheet("Summary")
                    ws_sum.set_column('B:C', 22)
                    
                    ws_sum.write('B2', '🏢 세관 구분', summary_header_format)
                    ws_sum.write('C2', '💰 수입부가세 총합계 (KRW)', summary_header_format)
                    
                    # 정산대장 시트의 데이터를 서머리하는 엑셀 수식 주입
                    last_row_idx = len(processed_data) + 1
                    ws_sum.write('B3', '안산세관', summary_data_format)
                    ws_sum.write_formula('C3', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "안산세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                    
                    ws_sum.write('B4', '안양세관', summary_data_format)
                    ws_sum.write_formula('C4', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "안양세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                    
                    ws_sum.write('B5', '부산세관', summary_data_format)
                    ws_sum.write_formula('C5', f'=SUMIF(통관월납_정산대장!L2:L{last_row_idx}, "부산세관", 통관월납_정산대장!D2:D{last_row_idx})', summary_data_format)
                    
                    ws_sum.write('B6', '🚀 전체 부가세 총계', summary_total_format)
                    ws_sum.write_formula('C6', '=SUM(C3:C5)', summary_total_format)
                    
                    # --- [시트 2 : 정산대장 메인 시트 생성] ---
                    df_final.to_excel(workbook, sheet_name="통관월납_정산대장", index=False)
                    ws = workbook.sheets["통관월납_정산대장"]
                    
                    # 기본 열너비 설정
                    ws.set_column('D:D', 18, num_format)
                    ws.set_column('G:G', 18, usd_format)
                    ws.set_column('I:I', 12, fx_format)
                    ws.set_column('J:K', 15, num_format)
                    
                    # 🛠️ [요청 사항 1 적용] 과세가격, 수식검증부가세, 검증결과 열너비 확장 (22~25 너비 확보)
                    ws.set_column('N:N', 24, num_format) # 과세가격
                    ws.set_column('O:O', 25, num_format) # 수식검증 부가세
                    ws.set_column('P:P', 22, align_center) # 고지액 검증 결과
                    
                    ws.write('N1', '과세가격(원화산출식)')
                    ws.write('O1', '수식검증 부가세(원단위 버림)')
                    ws.write('P1', '고지액 검증 결과')
                    
                    # 수식 주입 및 🛠️ [요청 사항 2 적용] 일치 심볼 주입 및 파란색 디자인 서식 연결
                    for i in range(2, len(processed_data) + 2):
                        ws.write_formula(f'N{i}', f'=(G{i}*I{i})+J{i}+K{i}', num_format)
                        ws.write_formula(f'O{i}', f'=ROUNDDOWN(N{i}*0.1, -1)', num_format)
                        
                        # IF 조건문 결과값을 심볼 매칭 형태로 수정
                        ws.write_formula(f'P{i}', f'=IF(D{i}=O{i}, "✔ 완벽 일치", "❌ 금액 불일치")', align_center)
                    
                    # 🎨 엑셀의 조건부 서식을 이용하여 P열에 '✔ 완벽 일치'가 뜨면 자동으로 파란색 글씨가 되도록 주입
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
                    
                    st.success("🎉 가독성 디자인 개선 및 세관별 Summary 연동이 모두 완료되었습니다!")
                    
                    st.download_button(
                        label="📥 최종 고도화 정산 마스터 대장 다운로드 (.xlsx)",
                        data=excel_data,
                        file_name="통관월납_정산_마스터대장_최종본_수식포함.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                    st.dataframe(df_final, use_container_width=True)
                    
                except Exception as e:
                    st.error(f"❌ 가독성 양식 업그레이드 중 오류가 발생했습니다: {e}")
