import streamlit as st
import pandas as pd
import docx
from io import BytesIO

# --- 페이지 설정 ---
st.set_page_config(
    page_title="성과관리 운영 현황 대시보드",
    layout="wide"
)

# --- 데이터 정의 (이미지 내용) ---
# 이 부분의 데이터를 수정하여 대시보드 내용을 변경할 수 있습니다.
title = "[작성요청] 성과관리 운영 현황 (예시적 / 각 사 상황에 맞게 기재)"

# 등급 배분
dist_method = "절대평가"
process_flow = "팀장 점수 평가/등급 제안 → 등급 심의 위원회 실시(위원장: 담당) → 개인별 등급 피드백(확정) → 이의 제기 → 최종 확정"

# 등급별 분포 현황 표 데이터
table_data = {
    '평가': ['S', 'A', 'B', 'C', 'D'],
    '책임(인원비중: 50%)': ['XX%', 'XX%', 'XX%', 'XX%', 'XX%'],
    '선임(인원비중: 30%)': ['XX%', 'XX%', 'XX%', 'XX%', 'XX%'],
    '사원(인원비중: 20%)': ['XX%', 'XX%', 'XX%', 'XX%', 'XX%'],
    'Total': ['XX%', 'XX%', 'XX%', 'XX%', 'XX%']
}
df = pd.DataFrame(table_data).set_index('평가').T

# 구성원 VOE
voe_list = [
    "🟢 절대평가 전환 이후 진급이나 연차에 따른 평가 왜곡은 많이 개선된 것으로 느껴짐",
    "🟢 절대평가 전환 이후 조직 내 경쟁이 줄어들어 동료 의식이 강화되고 팀워크에 도움이 되는 것 같음",
    "🟢 상대평가는 연초에 어떤 업무를 부여받는 지에 따라 평가가 결정되는 경우가 다수인데, 절대평가를 통해 이러한 부분이 많이 개선되었음",
    "🔴 절대평가 전환 이후 관대화가 많이 이뤄져 적당히 해도 B를 받는다고 느끼거나 B 평가를 수용하지 못함",
    "🔴 절대평가 전환 이후 재원은 한정된 상황에서 평가가 관대화되다 보니 전반적인 임금경쟁력이 낮아지는 경향이 있고, 고성과자에 대한 동기부여도 되지 않음. 차라리 상대평가가 나을 것 같음",
    "🔴 평가 기준이 조직별로 달라 타 팀의 A와 나 팀의 A가 같지 않은 경향이 있음"
]

# 평가 운영상의 Issue
issue_list = [
    "- Pay Band를 기준으로 평가/보상을 분리하지 않아 저 직급 인원의 저평가 현상이 나타남.",
    "- 평가 공정성 제고를 위해 수시 성과관리를 강화하려고 하나, 조직 책임자들의 업무 Load 증가로 인한 불만 증가",
    "- 성과 평가 보완을 위해 동료 평가를 도입하려고 하나, 구성원 수용성 부족",
    "- 역량과 성과를 보상에 연계하려고 하나, 역량 평가에 대한 공신력 부족",
    "- 이의제기 절차에 대한 평가 변경 / 유지에 대한 모호성",
    "- 보상과 평가 분리 요구",
    "(기타 각 사에서 성과관리 강화를 위해 개선이 필요한 사항들 / 타사 의견을 들어보고 싶은 사례들에 대해 기재)"
]


# --- Word 문서 생성 함수 ---
def create_word_document():
    doc = docx.Document()
    
    # 제목
    doc.add_heading(title, level=1)

    # 1. 등급 배분 섹션
    doc.add_heading('등급 배분', level=2)
    doc.add_paragraph(f"• 배분 방식: {dist_method}")
    doc.add_paragraph(f"• Process: {process_flow}")
    doc.add_paragraph() # 공백 추가

    # 2. 등급별 분포 현황 섹션
    doc.add_heading('등급별 분포 현황', level=2)
    # 테이블 추가
    table = doc.add_table(rows=df.shape[0] + 1, cols=df.shape[1] + 1)
    table.style = 'Table Grid'
    
    # 테이블 헤더 (첫 행)
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '직급'
    for i, col_name in enumerate(df.columns):
        hdr_cells[i+1].text = col_name

    # 테이블 바디
    for i, (index, row) in enumerate(df.iterrows()):
        row_cells = table.rows[i+1].cells
        row_cells[0].text = index
        for j, value in enumerate(row):
            row_cells[j+1].text = str(value)
    doc.add_paragraph()

    # 3. 구성원 VOE 섹션
    doc.add_heading('구성원 VOE', level=2)
    for item in voe_list:
        # 이모지를 텍스트로 추가
        p = doc.add_paragraph()
        p.add_run(item).font.name = 'Arial' # 이모지가 잘 보이도록 폰트 지정 (선택사항)
    doc.add_paragraph()
    
    # 4. 평가 운영상의 Issue 섹션
    doc.add_heading('평가 운영상의 Issue', level=2)
    for item in issue_list:
        doc.add_paragraph(item, style='List Bullet')
    
    # 메모리에 문서를 저장하여 BytesIO 객체로 반환
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- Streamlit UI 구성 ---
st.title(title)
st.markdown("---")

# 1. 등급 배분
with st.container(border=True):
    st.subheader("등급 배분")
    st.markdown(f"**배분 방식:** {dist_method}")
    st.markdown(f"**Process:** {process_flow}")

st.write("") # 간격

# 2. 등급별 분포 현황
cols = st.columns([0.2, 0.8], gap="medium")
with cols[0]:
    with st.container(border=True):
        st.subheader("등급별 분포 현황")
with cols[1]:
    st.table(df)

st.write("") # 간격

# 3. 구성원 VOE
cols = st.columns([0.2, 0.8], gap="medium")
with cols[0]:
    with st.container(border=True):
        st.subheader("구성원 VOE")
with cols[1]:
    for item in voe_list:
        st.markdown(item)

st.write("") # 간격

# 4. 평가 운영상의 Issue
cols = st.columns([0.2, 0.8], gap="medium")
with cols[0]:
    with st.container(border=True):
        st.subheader("평가 운영상의 Issue")
with cols[1]:
    for item in issue_list:
        st.markdown(item)

st.markdown("---")

# --- 다운로드 버튼 ---
st.write("### 보고서 다운로드")
word_file = create_word_document()
st.download_button(
    label="📥 Word 파일로 다운로드",
    data=word_file,
    file_name="성과관리_운영현황_보고서.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)
