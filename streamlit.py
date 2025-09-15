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

# ===== 전형별 코드 =====
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

# ===== 요청 헤더 / 쿠키 =====
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

# ===== Streamlit UI =====
st.set_page_config(layout="wide")
st.title("대학 입시자료 조회 및 다운로드")

# ===== GitHub에서 대학 목록 로드 =====
@st.cache_data(show_spinner=False)
def load_university_list(github_url):
    response = requests.get(github_url)
    response.raise_for_status()
    file_bytes = BytesIO(response.content)
    df = pd.read_excel(file_bytes, engine='openpyxl')
    df = df.dropna(subset=[df.columns[0], df.columns[1]])
    return df

# GitHub 파일 URL
GITHUB_URL = "https://raw.githubusercontent.com/사용자명/저장소명/브랜치/대학교별 코드.xlsx"

df = load_university_list(GITHUB_URL)
univ_list = df["학교명"].tolist()

# 사이드바
with st.sidebar:
    search_year = st.number_input("학년도 입력", min_value=2000, max_value=2100, value=SEARCH_YEAR_DEFAULT, step=1)
    selected_univ = st.selectbox("대학 선택", univ_list)
    types_options = ["전체"] + list(types_results.keys()) + list(types_main.keys())
    selected_type = st.selectbox("전형 선택", types_options)

# ===== 모집요강 PDF 다운로드 =====
@st.cache_data(show_spinner=False)
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

# ===== 전형별 입시자료 크롤링 =====
def crawl_admission_result_single(unv_cd, search_syr, sheet_name):
    if sheet_name in types_main:
        codes = types_main[sheet_name]
    elif sheet_name in types_results:
        codes = types_results[sheet_name]
    else:
        return None
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
            return combined_df
        else:
            return None
    except Exception as e:
        st.warning(f"{sheet_name} 크롤링 실패: {e}")
        return None

# ===== 버튼 클릭 후 크롤링 시작 =====
if st.button("크롤링 시작"):
    row = df[df["학교명"] == selected_univ].iloc[0]
    unv_cd = str(row["코드번호"]).zfill(7)

    st.info(f"{selected_univ} 입시자료 로딩 중... ⏳")

    # PDF 다운로드
    pdf_buffers = extract_and_download_pdfs(unv_cd, search_year, selected_univ)

    # ===== 오른쪽 화면 상/하 프레임 =====
    top_container = st.container()   # 주요사항
    bottom_container = st.container() # 입시결과

    # 상단: 주요사항
    st.subheader(f"📌 {search_year} 전형별 주요사항")
    for sheet_name in types_main.keys():
        if selected_type != "전체" and selected_type != sheet_name:
            continue
        placeholder = top_container.empty()
        df_sheet = crawl_admission_result_single(unv_cd, search_year, sheet_name)
        if df_sheet is not None:
            df_to_show = wrap_long_text(df_sheet, max_len=50)
            placeholder.markdown(f"**{sheet_name}**")
            placeholder.dataframe(df_to_show, use_container_width=True)

    # 하단: 입시결과
    st.subheader(f"📊 {search_year-1}학년도 입시결과")
    for sheet_name in types_results.keys():
        if selected_type != "전체" and selected_type != sheet_name:
            continue
        placeholder = bottom_container.empty()
        df_sheet = crawl_admission_result_single(unv_cd, search_year, sheet_name)
        if df_sheet is not None:
            placeholder.markdown(f"**{sheet_name}**")
            placeholder.dataframe(df_sheet, use_container_width=True)

    # Excel 다운로드
    excel_buffer = BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        for sheet_name in list(types_main.keys()) + list(types_results.keys()):
            if selected_type != "전체" and selected_type != sheet_name:
                continue
            df_sheet = crawl_admission_result_single(unv_cd, search_year, sheet_name)
            if df_sheet is not None:
                df_sheet.to_excel(writer, sheet_name=sheet_name[:31], index=False, header=False)
    excel_buffer.seek(0)
    st.download_button(
        label="📥 입시결과 다운로드",
        data=excel_buffer,
        file_name=f"{sanitize_filename(selected_univ)}_{search_year-1}년_대학입시결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # PDF 다운로드
    if pdf_buffers:
        st.markdown("### 모집요강 PDF 다운로드")
        for label, (content, fname) in pdf_buffers.items():
            st.download_button(
                label=f"📄 {label} 다운로드",
                data=content,
                file_name=fname,
                mime="application/pdf"
            )
    else:
        st.warning("모집요강 PDF가 없습니다.")

    st.success("크롤링 완료! ✅")
