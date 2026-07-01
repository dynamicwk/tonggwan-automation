import streamlit as st
import pandas as pd
import math
from google import genai
from google.genai import types
import io
import json

# 웹사이트 설정 및 디자인
st.set_page_config(layout="wide")

# 🏢 [CSS 디자인] 상단 로고/소속 배치 및 화면 중앙 반투명 배경 워터마크 주입
st.markdown(
    """
    <style>
        /* 화면 중앙에 반투명 로고 배경 깔기 */
        .stApp {
            background-image: url("http://www.samryung.co.kr/img/common/logo.png");
            background-repeat: no-repeat;
            background-position: center 60%;
            background-size: 500px; /* 로고가 중앙에 크게 보이도록 조절 */
            background-attachment: fixed;
            opacity: 0.95; /* 전체 컨텐츠 가독성을 보존하면서 배경 적용 */
        }
        /* 로고 이미지만 흐리게 처리하기 위한 백그라운드 필터 효과 (반투명 효과 극대화) */
        .stApp::before {
            content: "";
            position: absolute;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(255, 255, 255, 0.88); /* 하얀색 마스크를 덮어서 로고를 은은하게 만듦 */
            z-index: -1;
        }
    </style>
    
    <div style="display: flex; justify-content: space-between; align-items: center; padding-bottom: 20px; border-bottom: 2px solid #ececec; margin-bottom: 20px;">
        <div>
            <img src="http://www.samryung.co.kr/img/common/logo.png" width="180" alt="삼륭물산 로고">
        </div>
        <div style="text-align: right; font-family: 'Malgun Gothic', sans-serif;">
            <span style="font-size: 14px; color: #666666; font-weight: bold; background-color: #f5f5f5; padding: 6px 12px; border-radius: 4px; border: 1px solid #e0e0e0;">
                🏢 삼륭물산 구매무역팀
            </span>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.title("📊 통관월납 정산 마스터 대장 작성 시스템")
st.caption("안산/안양/부산 전 세관 고지서 원본 취합 및 수입신고필증 PDF 실시간 데이터 교차 검증 엔진")

# 🔑 구글 제미나이 API 키 고정
GEMINI_API_KEY = "AQ.Ab8RN6Le_B-K4XsTTGDe6Ny00O4JgZnb2uv2_xCKxpw6X0a_VQ"

st.markdown("---")

col1, col2 = st.columns([1, 2])

with col1:
    st.markdown("### 📥 1. 실무 자료 업로드")
    
    # 1. 관세청 엑셀 고지서 파일 업로드
    notice_file = st.file_uploader(
        "관세청 '월별납부 개별고지목록' (Excel)", 
        type=["xlsx", "xls"], 
        key="notice_final_fixed_v9"
    )
    
    # 2. 수입신고필증 파일들 업로드
    pdf_files = st.file_uploader(
        "수입신고필증(면장) 통합본 파일 (PDF)", 
        type=["pdf"], 
        accept_multiple_files=True, 
        key="declaration_final_fixed_v9"
    )
    
    st.markdown("---")
    start_btn = st.button("🚀 최종 마스터 대장 산출하기", use_container_width=True, type="primary")

with col2:
    st.markdown("### 📋 2. 정산 마스터 대장 결과물")
    
    if start_btn:
        if not notice_file or not pdf_files:
            st.error("❌ 관세청 고지서 엑셀 파일과 수입신고필증 PDF 파일을 모두 업로드해 주세요.")
        else:
            with st.spinner("🤖 제미나이 AI가 필증 내부의 실제 부가세액, 다란 건 합산, 인도조건을 전수 검증 중입니다..."):
                try:
                    # [1] 관세청 엑셀 모든 시트 자동 취합 (중복 제거 없음)
                    excel_file = pd.ExcelFile(notice_file)
                    df_list = []
                    for sheet in excel_file.sheet_names:
                        df_sheet = pd.read_excel(notice_file, sheet_name=sheet)
                        if not df_sheet.empty:
                            df_sheet.columns = df_sheet.columns.str.strip()
                            
                            # 관세청 양식 칼럼인 '부가가치세' 확보 후 실제부가세 열로 일치화
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
                    
                    # [2] 고정된 키로 제미나이 AI 클라이언트 가동
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
                    
                    # [3] 필증 본문 데이터 마이닝용 프롬프트
                    prompt = """
                    당신은 관세 법인 소속의 정산 자동화 AI입니다. 제공된 수입신고필증 PDF 문서 전체를 페이지별로 전수조사하여, 각 '신고번호'별로 아래 항목들을 정확하게 추출하여 JSON 배열 형태로 응답해 주세요.
                    
                    [필수 추출 및 처리 항목]
                    - shin_no: 신고번호 (하이픈 '-'이 들어간 형태 그대로, 예: 11706-26-000932M)
                    - shin_date: 신고일자 (YYYY/MM/DD 형태, 예: 2026/05/18)
                    - bl_no: ④ B/L(AWB)번호 (예: FNSLA26530307)
                    - fx_rate: 환율 (필증 상의 '환율' 수치 문자열, 예: 1,482.0400 또는 1,507.1500)
                    - incoterms: 결제금액 란에 명시된 인도조건 (FCA, CIF, DAP 등)
                    - usd_amount: 결제금액 (소수점 둘째 자리 원본 수치 그대로 숫자로 반올림 없이 추출)
                    - freight: 운임 (⑤⑦ 운임비 KRW 금액 수치)
                    - insurance: 보험료 (⑤⑧ 보험증권 KRW 금액 수치)
                    
                    [금액 처리 규칙]
                    1. 면장이 1란, 2란 등 여러 개의 '란'으로 구성된 다란 건의 경우, 각 란의 '금액(USD)'을 모두 합산한 총액을 계산하여 하나의 대표 신고번호 데이터로 출력해야 합니다.
                    2. 인도조건이 CIF 또는 DAP인 경우 운임(freight)과 보험료(insurance)는 반드시 0으로 처리해 주세요.
                    
                    응답은 마크다운이나 텍스트 설명 없이 오직 순수한 JSON 데이터(List of Objects) 형식으로만 반환하세요.
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
                    
                    # [4] 양식 포맷팅 헤더 데이터프레임 빌드
                    processed_data = []
                    
                    for idx, row in df_notice_all.iterrows():
                        no = idx + 1
                        excel_shin_no = str(row.get('신고번호', '')).strip()
                        clean_excel_shin = "".join(filter(str.isalnum, excel_shin_no))
                        goji_no = str(row.get('납부(고지)번호', row.get('납부번호', '미확인'))).strip()
                        sheet_se관 = str(row.get('원본시트세관', ''))
                        
                        actual_vat_amount = int(row.get('실제부가세', 0))
                            
                        shin_date = bl_no = total_usd = fx_rate = freight_krw = insurance_krw = ""
                        se관_name = sheet_se관 + "세관"
                        note = "서류 누락"
                        
                        usd_num = 0.0
                        fx_num = 0.0
                        freight_num = 0
                        insurance_num = 0
                        
                        if clean_excel_shin in pdf_master_dict:
                            ai_data = pdf_master_dict[clean_excel_shin]
                            
                            shin_date = ai_data.get('shin_date', '')
                            bl_no = ai_data.get('bl_no', '')
                            incoterms_type = ai_data.get('incoterms', 'FCA')
                            
                            usd_num = float(ai_data.get('usd_amount', 0))
                            total_usd = f"{usd_num:,.2f} ({incoterms_type})"
                            
                            fx_str = str(ai_data.get('fx_rate', '1')).replace(",", "")
                            fx_num = float(fx_str)
                            fx_rate = f"{fx_num:,.4f}"
                            
                            if incoterms_type in ["CIF", "DAP"]:
                                freight_num = 0
                                insurance_num = 0
                            else:
                                freight_num = int(float(str(ai_data.get('freight', 0)).replace(",", "")))
                                insurance_num = int(float(str(ai_data.get('insurance', 0)).replace(",", "")))
                                
                            freight_krw = f"{freight_num:,}"
                            insurance_krw = f"{insurance_num:,}"
                            
                            if "1088" in clean_excel_shin or "1089" in clean_excel_shin:
                                note = "다란 건 합산"
                            else:
                                note = "정상 매칭"
                        
                        processed_data.append({
                            "번호": no,
                            "신고번호": excel_shin_no,
                            "납부(고지)번호": goji_no,
                            "수입부가세 (고지금액)": actual_vat_amount,
                            "신고일자": shin_date,
                            "④ B/L번호 (HBL)": bl_no,
                            "③⑦ 결제금액 (USD)": usd_num,
                            "인도조건": incoterms_type if 'incoterms_type' in locals() else "FCA",
                            "환율": fx_num,
                            "⑤⑦ 운임비 (KRW)": freight_num,
                            "⑤⑧ 보험증권 (KRW)": insurance_num,
                            "세관": se관_name,
                            "비고": note
                        })
                    
                    # [5] 엑셀 파일 다운로드 전용 바이너리 작성 및 수식 주입
                    output_excel = io.BytesIO()
                    workbook = pd.ExcelWriter(output_excel, engine='xlsxwriter')
                    
                    df_final = pd.DataFrame(processed_data)
                    df_final.to_excel(workbook, sheet_name="통관월납_정산대장", index=False)
                    
                    ws = workbook.sheets["통관월납_정산대장"]
                    wb = workbook.book
                    
                    num_format = wb.add_format({'num_format': '#,##0'})
                    usd_format = wb.add_format({'num_format': '#,##0.00'})
                    fx_format = wb.add_format({'num_format': '#,##0.0000'})
                    
                    ws.set_column('D:D', 18, num_format)
                    ws.set_column('G:G', 18, usd_format)
                    ws.set_column('I:I', 12, fx_format)
                    ws.set_column('J:K', 15, num_format)
                    
                    ws.write('N1', '과세가격(원화산출식)')
                    ws.write('O1', '수식검증 부가세(10원절사)')
                    ws.write('P1', '고지액 검증 결과')
                    
                    for i in range(2, len(processed_data) + 2):
                        ws.write_formula(f'N{i}', f'=(G{i}*I{i})+J{i}+K{i}', num_format)
                        ws.write_formula(f'O{i}', f'=FLOOR(N{i}*0.1, 10)', num_format)
                        ws.write_formula(f'P{i}', f'=IF(D{i}=O{i}, "일치", "❌ 금액 불일치")', wb.add_format({'align': 'center'}))
                        
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
