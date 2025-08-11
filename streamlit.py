import streamlit as st
import pandas as pd
import requests
from bs4 import BeautifulSoup
import time
from io import BytesIO

# --- 전형 코드 매핑 ---
types = {
    "학생부종합": {"upcd": "20", "cd": "22"},
    "학생부교과": {"upcd": "30", "cd": "32"},
    "수능": {"upcd": "40", "cd": "42"},
}

# --- 대학 목록 불러오기 (GitHub URL 사용) ---
@st.cache_data
def load_university_list(github_url):
    try:
        response = requests.get(github_url)
        response.raise_for_status()
        file_bytes = BytesIO(response.content)
        df = pd.read_excel(file_bytes, engine='openpyxl')
        df = df.dropna(subset=[df.columns[0], df.columns[1]])
        return df.values.tolist()
    except requests.exceptions.RequestException as e:
        st.error(f"GitHub에서 파일을 불러오는 데 실패했습니다: {e}")
        return None

# --- 크롤링 함수: 여러 테이블 처리 로직 ---
def crawl_admission_result(univ_name, univ_code, selected_types):
    cookies = {
        'WMONID': 'NYfDEAkX3Jy',
        'JSESSIONID': 'V9Tor4qz9JI1R0wOWXqKXhcJbeLiyXWdTSgfWj1hzo1aRGzUlCTAoSQSWOuxxFFK.amV1c19kb21haW4vYWRpZ2Ex',
    }
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'ko-KR,ko;q=0.9',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Origin': 'https://www.adiga.kr',
        'Referer': 'https://www.adiga.kr/uct/acd/ade/criteriaAndResultPopup.do',
        'User-Agent': 'Mozilla/5.0',
        'X-CSRF-TOKEN': 'b4561457-4e76-449b-9099-c36118c3f560',
        'X-Requested-With': 'XMLHttpRequest',
    }

    sheet_data = {}

    for name in selected_types:
        all_data = []
        codes = types.get(name)
        if not codes:
            continue
        data = {
            '_csrf': 'b4561457-4e76-449b-9099-c36118c3f560',
            'searchSyr': '2026',
            'unvCd': str(univ_code).zfill(7),
            'searchUnvComp': '0',
            'tsrdCmphSlcnArtclUpCd': codes['upcd'],
            'tsrdCmphSlcnArtclCd': codes['cd'],
        }
        response = requests.post(
            'https://www.adiga.kr/uct/acd/ade/criteriaAndResultItemAjax.do',
            cookies=cookies,
            headers=headers,
            data=data,
        )
        time.sleep(0.5)

        soup = BeautifulSoup(response.text, 'html.parser')
        tables = soup.select("table")
        if not tables:
            continue

        for table in tables:
            span_map = {}
            table_matrix = []
            rows = table.find_all('tr')
            for row_idx, row in enumerate(rows):
                cells = row.find_all(['td', 'th'])
                current_row = [None] * 100
                col_idx = 0
                for cell in cells:
                    while (row_idx, col_idx) in span_map:
                        current_row[col_idx] = span_map[(row_idx, col_idx)][2]
                        col_idx += 1
                    text = cell.get_text(strip=True)
                    rowspan = int(cell.get('rowspan', 1))
                    colspan = int(cell.get('colspan', 1))
                    for r in range(rowspan):
                        for c in range(colspan):
                            if r == 0 and c == 0:
                                current_row[col_idx] = text
                            else:
                                span_map[(row_idx + r, col_idx + c)] = (rowspan, colspan, text)
                    col_idx += colspan
                table_matrix.append(current_row[:col_idx])
            df = pd.DataFrame(table_matrix).fillna('')
            all_data.append(df)
            all_data.append(pd.DataFrame([['' for _ in range(df.shape[1])]]))

        if all_data:
            combined = pd.concat(all_data, ignore_index=True)
            sheet_data[name] = combined

    return sheet_data

# --- Streamlit 앱 시작 ---
st.title("🎓 2025 대학 입시 결과 크롤링")

GITHUB_RAW_URL = "https://raw.githubusercontent.com/fresh601/university_crawlig/main/대학교별%20코드.xlsx"
univ_list = load_university_list(GITHUB_RAW_URL)

if univ_list is not None:
    univ_dict = {name: code for name, code in univ_list}

    selected_univ = st.selectbox("🏫 대학 선택", list(univ_dict.keys()))
    selected_types = st.multiselect("📌 전형 선택", list(types.keys()), default=["학생부종합"])

    if st.button("📊 입시결과 불러오기"):
        with st.spinner(f"{selected_univ}의 입시결과를 가져오는 중입니다..."):
            sheet_data = crawl_admission_result(selected_univ, univ_dict[selected_univ], selected_types)

            if not sheet_data:
                st.warning("해당 대학의 선택한 전형 결과가 없습니다.")
            else:
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    for sheet_name, df in sheet_data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                st.success("🎉 크롤링 완료!")
                st.download_button(
                    label="📥 엑셀 파일 다운로드",
                    data=output.getvalue(),
                    file_name=f"{selected_univ}_2025년_대학입시결과.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                for sheet_name, df in sheet_data.items():
                    st.subheader(f"📄 {sheet_name}")
                    st.dataframe(df)
else:
    st.info("GitHub에서 '대학교별 코드.xlsx' 파일을 불러오는 데 실패했습니다. URL을 확인해 주세요.")



