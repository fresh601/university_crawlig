import os
import re
import time
import unicodedata
import requests
import pandas as pd
from bs4 import BeautifulSoup
from io import BytesIO, StringIO
import streamlit as st

# ===== 고정 설정 =====
BASE = "https://www.adiga.kr"
DETAIL_URL = f"{BASE}/ucp/uvt/uni/univDetail.do"
DOWNLOAD_URL = f"{BASE}/cmm/com/file/fileDown.do"
MENU_ID = "PCUVTINF2000"
SEARCH_YEAR_DEFAULT = 2026
UNIV_LIST_PATH = "대학교별 코드.xlsx"  # 깃허브에 포함될 파일

# ===== 유틸 함수 =====
def sanitize_filename(name: str) -> str:
    name = unicodedata.normalize("NFKC", str(name))
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def norm_text(el) -> str:
    return ' '.join(el.get_text(separator=' ', strip=True).split())

def wrap_long_text(df, max_len=50):
    df_wrapped = df.copy()
    for col in df_wrapped.columns:
        df_wrapped[col] = df_wrapped[col].apply(
            lambda x: "\n".join([str(x)[i:i+max_len] for i in range(0, len(str(x)), max_len)])
        )
    return df_wrapped

# ===== 입시결과 크롤링 =====
cookies = {
    'WMONID': 'NYfDEAkX3Jy',
    'JSESSIONID': 'V9Tor4qz9JI1R0wOWXqKXhcJbeLiyXWdTSgfWj1hzo1aRGbUlCTAoSQSWOuxxFFK.amV1c19kb21haW4vYWRpZ2Ex',
}
headers = {
    'Accept': 'application/json, text/plain, */*',
    'Content-Type': 'application/x-www-form-urlencoded',
    'Origin': 'https://www.adiga.kr',
    'Referer': 'https://www.adiga.kr/uct/acd/ade/criteriaAndResultPopup.do',
    'User-Agent': 'Mozilla/5.0',
    'X-CSRF-TOKEN': 'b4561457-4e76-449b-909b-9099-c36118c3f560',
    'X-Requested-With': 'XMLHttpRequest',
}

types_results = {
    "학생부종합": {"upcd": "20", "cd": "22"},
    "학생부교과": {"upcd": "30", "cd": "32"},
    "수능": {"upcd": "40", "cd": "42"},
}
types_main = {
    "학생부종합(주요사항)": {"upcd": "20", "cd": "21"},
    "학생부교과(주요사항)": {"upcd": "30", "cd": "31"},
    "수능(주요사항)": {"upcd": "40", "cd": "41"},
}

def crawl_admission_results_chunk(unv_cd, search_syr, name, codes):
    sheet_data = {}
    data = {
        '_csrf': headers['X-CSRF-TOKEN'],
        'searchSyr': search_syr,
        'unvCd': str(unv_cd).zfill(7),
        'compUnvCd': '',
        'searchUnvComp': '0',
        'tsrdCmphSlcnArtclUpCd': codes['upcd'],
        'tsrdCmphSlcnArtclCd': codes['cd'],
    }
    try:
        response = requests.post(
            'https://www.adiga.kr/uct/acd/ade/criteriaAndResultItemAjax.do',
            cookies=cookies, headers=headers, data=data, timeout=30
        )
        time.sleep(0.2)
        soup = BeautifulSoup(response.text, 'lxml')
        tables = soup.find_all('table')
        df_list = []
        for table in tables:
            try:
                df_table = pd.read_html(StringIO(str(table)), flavor='lxml')[0]
                df_list.append(df_table)
                df_list.append(pd.DataFrame([['' for _ in range(df_table.shape[1])]]))
            except:
                continue
        if df_list:
            combined_df = pd.concat(df_list, ignore_index=True)
            sheet_data[name] = combined_df
    except Exception as e:
        st.warning(f"{name} 크롤링 실패: {e}")
    return sheet_data

def extract_and_download_pdfs(unv_cd, search_syr, univ_name):
    plan_ids = susi_ids = jeongsi_ids = None
    params = {"menuId": MENU_ID, "unvCd": unv_cd, "searchSyr": search_syr}
    headers_req = {"User-Agent": "Mozilla/5.0"}
    res = requests.get(DETAIL_URL, params=params, headers=headers_req, timeout=30)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")
    ul = soup.select_one("ul#fileResult")
    if ul:
        for li in ul.select("li"):
            a = li.select_one("a[onclick]")
            span = li.select_one("span")
            if not a or not span:
                continue
            text = norm_text(span)
            onclick = a.get("onclick", "")
            m = re.search(r"fnUnvFileDownOne\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,", onclick)
            if not m:
                continue
            file_id, file_sn = m.group(1), m.group(2)
            if ("대학입학전형" in text) and ("시행계획" in text):
                plan_ids = (file_id, file_sn)
            elif ("수시" in text) and ("모집요강" in text):
                susi_ids = (file_id, file_sn)
            elif ("정시" in text) and ("모집요강" in text):
                jeongsi_ids = (file_id, file_sn)
    pdf_buffers = {}
    for label, ids in [("시행계획", plan_ids), ("수시", susi_ids), ("정시", jeongsi_ids)]:
        if ids:
            f_id, f_sn = ids
            params_file = {
                "fileId": f_id,
                "fileSn": f_sn,
                "menuId": MENU_ID,
                "downLogYn": "Y",
                "unvCd": unv_cd,
                "searchSyr": search_syr,
                "_": str(int(time.time() * 1000)),
            }
            headers_file = {
                "User-Agent": "Mozilla/5.0",
                "Referer": f"{DETAIL_URL}?menuId={MENU_ID}&unvCd={unv_cd}&searchSyr={search_syr}",
                "X-Requested-With": "XMLHttpRequest",
            }
            r = requests.get(DOWNLOAD_URL, params=params_file, headers=headers_file, timeout=60)
            if r.status_code == 200:
                fname = sanitize_filename(f"{univ_name}_{label}_모집요강.pdf")
                pdf_buffers[label] = (r.content, fname)
    return pdf_buffers

# ===== Streamlit UI =====
st.set_page_config(layout="wide")
st.title("대학 입시자료 조회 및 다운로드")

if not os.path.exists(UNIV_LIST_PATH):
    st.error(f"{UNIV_LIST_PATH} 파일이 없습니다. 깃허브에 포함시켜주세요.")
else:
    df = pd.read_excel(UNIV_LIST_PATH)
    if "코드번호" not in df.columns or "학교명" not in df.columns:
        st.error("'코드번호'와 '학교명' 열이 필요합니다.")
    else:
        univ_list = df["학교명"].tolist()

        # 사이드바
        with st.sidebar:
            search_year = st.number_input("학년도 입력", min_value=2000, max_value=2100,
                                          value=SEARCH_YEAR_DEFAULT, step=1)
            selected_univ = st.selectbox("대학 선택", univ_list)
            types_options = ["전체"] + list(types_results.keys()) + list(types_main.keys())
            selected_type = st.selectbox("전형 선택", types_options)

        # ===== Placeholder 준비 =====
        top_container = st.container()   # 주요사항
        bottom_container = st.container() # 입시결과
        pdf_container = st.container()    # PDF 다운로드
        status_placeholder = st.empty()   # 상태 메시지 placeholder
        progress_bar = st.progress(0)

        if st.button("크롤링 시작") or "admission_data" in st.session_state:

            if "admission_data" not in st.session_state:
                row = df[df["학교명"] == selected_univ].iloc[0]
                unv_cd = str(row["코드번호"]).zfill(7)
                st.session_state.admission_data = {}
                st.session_state.pdf_buffers = {}

                all_types = {**types_results, **types_main}
                total = len(all_types)
                for i, (name, codes) in enumerate(all_types.items(), 1):
                    status_placeholder.info(f"{name} 크롤링 중... ({i}/{total})")
                    data_chunk = crawl_admission_results_chunk(unv_cd, search_year, name, codes)
                    st.session_state.admission_data.update(data_chunk)

                    progress_bar.progress(i / total)

                status_placeholder.info("PDF 크롤링 중...")
                st.session_state.pdf_buffers = extract_and_download_pdfs(unv_cd, search_year, selected_univ)

            # ===== 화면 표시 =====
            with top_container:
                if any("주요사항" in name for name in st.session_state.admission_data.keys()):
                    st.header("주요사항")
                    for sheet_name, df_sheet in st.session_state.admission_data.items():
                        if "주요사항" not in sheet_name:
                            continue
                        if selected_type != "전체" and selected_type != sheet_name:
                            continue
                        st.markdown(f"**{sheet_name}**")
                        st.dataframe(wrap_long_text(df_sheet, max_len=50), use_container_width=True)

            with bottom_container:
                if any("주요사항" not in name for name in st.session_state.admission_data.keys()):
                    st.header("입시결과")
                    for sheet_name, df_sheet in st.session_state.admission_data.items():
                        if "주요사항" in sheet_name:
                            continue
                        if selected_type != "전체" and selected_type != sheet_name:
                            continue
                        st.markdown(f"**{sheet_name}**")
                        st.dataframe(wrap_long_text(df_sheet, max_len=50), use_container_width=True)

                    # ===== 입시결과 Excel 다운로드 버튼 =====
                    excel_buffer = BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                        for sheet_name, df_sheet in st.session_state.admission_data.items():
                            if selected_type != "전체" and selected_type != sheet_name:
                                continue
                            df_sheet.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)
                    excel_buffer.seek(0)
                    st.download_button(
                        label="📥 입시결과 다운로드",
                        data=excel_buffer,
                        file_name=f"{sanitize_filename(selected_univ)}_{search_year-1}년_대학입시결과.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            # ===== PDF 다운로드 =====
            with pdf_container:
                if st.session_state.pdf_buffers:
                    st.markdown("### 모집요강 PDF 다운로드")
                    for label, (content, fname) in st.session_state.pdf_buffers.items():
                        st.download_button(
                            label=f"📄 {label} 다운로드",
                            data=content,
                            file_name=fname,
                            mime="application/pdf"
                        )
                else:
                    st.warning("모집요강 PDF가 없습니다.")

            status_placeholder.success("크롤링 완료! ✅")
