import os
import re
import time
import unicodedata
import requests
import pandas as pd

from bs4 import BeautifulSoup
from io import BytesIO, StringIO

import streamlit as st


# ============================================================
# 기본 설정
# ============================================================

BASE = "https://www.adiga.kr"

DETAIL_URL = (
    f"{BASE}/ucp/uvt/uni/univDetail.do"
)

DOWNLOAD_URL = (
    f"{BASE}/cmm/com/file/fileDown.do"
)

POPUP_URL = (
    f"{BASE}/uct/acd/ade/criteriaAndResultPopup.do"
)

# 기존 입시결과 AJAX
OLD_RESULT_URL = (
    f"{BASE}/uct/acd/ade/criteriaAndResultItemAjax.do"
)

# 2027학년도부터 변경된 입시결과 AJAX
NEW_RESULT_URL = (
    f"{BASE}/uct/acd/ade/criteriaAndResultItemNewAjax.do"
)

MENU_ID = "PCUVTINF2000"

SEARCH_YEAR_DEFAULT = 2027

UNIV_LIST_PATH = "대학교별 코드.xlsx"


# ============================================================
# 유틸 함수
# ============================================================

def sanitize_filename(name: str) -> str:

    name = unicodedata.normalize(
        "NFKC",
        str(name)
    )

    name = re.sub(
        r'[<>:"/\\|?*\x00-\x1F]',
        " ",
        name
    )

    name = re.sub(
        r"\s+",
        " ",
        name
    ).strip()

    return name


def norm_text(el) -> str:

    return " ".join(
        el.get_text(
            separator=" ",
            strip=True
        ).split()
    )


def wrap_long_text(
    df,
    max_len=50
):

    df_wrapped = df.copy()

    for col in df_wrapped.columns:

        df_wrapped[col] = df_wrapped[col].apply(

            lambda x:
                "\n".join(
                    [
                        str(x)[i:i + max_len]
                        for i in range(
                            0,
                            len(str(x)),
                            max_len
                        )
                    ]
                )

        )

    return df_wrapped


# ============================================================
# 세션 생성
# ============================================================

def create_session():

    session = requests.Session()

    session.headers.update({

        "User-Agent":
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/153.0.0.0 Safari/537.36",

        "Accept-Language":
            "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7"

    })

    return session


# ============================================================
# CSRF 토큰 가져오기
# ============================================================

def get_csrf_token(
    session,
    unv_cd,
    search_syr
):

    params = {

        "searchSyr":
            str(search_syr),

        "unvCd":
            str(unv_cd).zfill(7),

        "tsrdCmphSlcnArtclUpCd":
            "20"

    }

    try:

        response = session.get(

            POPUP_URL,

            params=params,

            timeout=30,

            verify=False

        )

        response.raise_for_status()

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        csrf_input = soup.select_one(
            'input[name="_csrf"]'
        )

        if csrf_input:

            return csrf_input.get(
                "value",
                ""
            )

    except Exception:

        pass

    return ""


# ============================================================
# 전형 구분
# ============================================================

types_results = {

    "학생부종합": {
        "upcd": "20",
        "cd": "22"
    },

    "학생부교과": {
        "upcd": "30",
        "cd": "32"
    },

    "수능": {
        "upcd": "40",
        "cd": "42"
    }

}


types_main = {

    "학생부종합(주요사항)": {
        "upcd": "20",
        "cd": "21"
    },

    "학생부교과(주요사항)": {
        "upcd": "30",
        "cd": "31"
    },

    "수능(주요사항)": {
        "upcd": "40",
        "cd": "41"
    }

}


# ============================================================
# 주요사항 크롤링
#
# 기존 AJAX 방식 유지
# ============================================================

def crawl_major_items(
    unv_cd,
    search_syr,
    name,
    codes
):

    sheet_data = {}

    session = create_session()

    csrf_token = get_csrf_token(
        session,
        unv_cd,
        search_syr
    )

    headers = {

        "Accept":
            "application/json, text/plain, */*",

        "Content-Type":
            "application/x-www-form-urlencoded",

        "Origin":
            BASE,

        "Referer":
            POPUP_URL,

        "User-Agent":
            "Mozilla/5.0",

        "X-Requested-With":
            "XMLHttpRequest"

    }

    if csrf_token:

        headers[
            "X-CSRF-TOKEN"
        ] = csrf_token


    data = {

        "_csrf":
            csrf_token,

        "searchSyr":
            str(search_syr),

        "unvCd":
            str(unv_cd).zfill(7),

        "compUnvCd":
            "",

        "searchUnvComp":
            "0",

        "tsrdCmphSlcnArtclUpCd":
            codes["upcd"],

        "tsrdCmphSlcnArtclCd":
            codes["cd"]

    }


    try:

        response = session.post(

            OLD_RESULT_URL,

            headers=headers,

            data=data,

            timeout=30,

            verify=False

        )

        response.raise_for_status()

    except Exception as e:

        st.warning(
            f"{name} 크롤링 실패: {e}"
        )

        return sheet_data


    time.sleep(0.2)


    try:

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

        tables = soup.find_all(
            "table"
        )

        df_list = []


        for table in tables:

            try:

                dfs = pd.read_html(
                    StringIO(
                        str(table)
                    )
                )

                if not dfs:
                    continue

                df_table = dfs[0]

                df_list.append(
                    df_table
                )

                # 표 사이 빈 줄
                df_list.append(

                    pd.DataFrame(

                        [
                            [
                                ""
                                for _ in range(
                                    df_table.shape[1]
                                )
                            ]
                        ]

                    )

                )

            except Exception:

                continue


        if df_list:

            combined_df = pd.concat(
                df_list,
                ignore_index=True
            )

            sheet_data[name] = (
                combined_df
            )


    except Exception as e:

        st.warning(
            f"{name} 표 변환 실패: {e}"
        )


    return sheet_data


# ============================================================
# 입시결과 크롤링
#
# 2026 이하 → 기존 AJAX
# 2027 이상 → 새로운 AJAX
# ============================================================

def crawl_admission_results_chunk(
    unv_cd,
    search_syr,
    name,
    codes
):

    sheet_data = {}

    search_syr = int(
        search_syr
    )


    # ========================================================
    # 2027학년도부터 새로운 AJAX
    # ========================================================

    if search_syr >= 2027:

        session = create_session()

        try:

            csrf_token = get_csrf_token(
                session,
                unv_cd,
                search_syr
            )

            data = {

                "searchSyr":
                    str(search_syr),

                "unvCd":
                    str(unv_cd).zfill(7),

                "tsrdCmphSlcnArtclUpCd":
                    codes["upcd"],

                "compUnvCd":
                    ""

            }


            headers = {

                "Accept":
                    "application/json, text/plain, */*",

                "Content-Type":
                    "application/x-www-form-urlencoded",

                "Origin":
                    BASE,

                "Referer":
                    POPUP_URL,

                "User-Agent":
                    "Mozilla/5.0",

                "X-Requested-With":
                    "XMLHttpRequest"

            }


            if csrf_token:

                headers[
                    "X-CSRF-TOKEN"
                ] = csrf_token


            response = session.post(

                NEW_RESULT_URL,

                headers=headers,

                data=data,

                timeout=30,

                verify=False

            )

            response.raise_for_status()

            html = response.text


        except Exception as e:

            st.warning(
                f"{name} 크롤링 실패: {e}"
            )

            return sheet_data


    # ========================================================
    # 2026학년도 이하
    #
    # 기존 코드 그대로 유지
    # ========================================================

    else:

        cookies = {

            "WMONID":
                "NYfDEAkX3Jy",

            "JSESSIONID":
                "V9Tor4qz9JI1R0wOWXqKXhcJbeLiyXWdTSgfWj1hzo1aRGbUlCTAoSQSWOuxxFFK.amV1c19kb21haW4vYWRpZ2Ex"

        }


        headers = {

            "Accept":
                "application/json, text/plain, */*",

            "Content-Type":
                "application/x-www-form-urlencoded",

            "Origin":
                BASE,

            "Referer":
                POPUP_URL,

            "User-Agent":
                "Mozilla/5.0",

            "X-CSRF-TOKEN":
                "b4561457-4e76-449b-909b-9099-c36118c3f560",

            "X-Requested-With":
                "XMLHttpRequest"

        }


        data = {

            "_csrf":
                headers["X-CSRF-TOKEN"],

            "searchSyr":
                search_syr,

            "unvCd":
                str(unv_cd).zfill(7),

            "compUnvCd":
                "",

            "searchUnvComp":
                "0",

            "tsrdCmphSlcnArtclUpCd":
                codes["upcd"],

            "tsrdCmphSlcnArtclCd":
                codes["cd"]

        }


        try:

            response = requests.post(

                OLD_RESULT_URL,

                cookies=cookies,

                headers=headers,

                data=data,

                timeout=30

            )

            response.raise_for_status()

            html = response.text


        except Exception as e:

            st.warning(
                f"{name} 크롤링 실패: {e}"
            )

            return sheet_data


    # ========================================================
    # HTML 표 파싱
    # ========================================================

    time.sleep(0.2)


    try:

        soup = BeautifulSoup(
            html,
            "html.parser"
        )

        tables = soup.find_all(
            "table"
        )

        df_list = []


        for table in tables:

            try:

                dfs = pd.read_html(
                    StringIO(
                        str(table)
                    )
                )

                if not dfs:
                    continue

                df_table = dfs[0]

                df_list.append(
                    df_table
                )

                df_list.append(

                    pd.DataFrame(

                        [
                            [
                                ""
                                for _ in range(
                                    df_table.shape[1]
                                )
                            ]
                        ]

                    )

                )

            except Exception:

                continue


        if df_list:

            combined_df = pd.concat(
                df_list,
                ignore_index=True
            )

            sheet_data[name] = (
                combined_df
            )


    except Exception as e:

        st.warning(
            f"{name} 표 변환 실패: {e}"
        )


    return sheet_data


# ============================================================
# 모집요강 파일 다운로드
#
# 2026 / 2027 공통
# ============================================================

def extract_and_download_files(
    unv_cd,
    search_syr,
    univ_name
):

    plan_ids = None
    susi_ids = None
    jeongsi_ids = None


    search_syr = str(
        search_syr
    )


    session = create_session()


    # ========================================================
    # 대학 상세 페이지
    # ========================================================

    params = {

        "menuId":
            MENU_ID,

        "unvCd":
            str(unv_cd).zfill(7),

        "searchSyr":
            search_syr

    }


    try:

        res = session.get(

            DETAIL_URL,

            params=params,

            timeout=30,

            verify=False

        )

        res.raise_for_status()

    except Exception as e:

        st.warning(
            f"모집요강 페이지 접속 실패: {e}"
        )

        return {}


    soup = BeautifulSoup(
        res.text,
        "html.parser"
    )


    ul = soup.select_one(
        "ul#fileResult"
    )


    if not ul:

        return {}


    # ========================================================
    # 파일 목록 탐색
    # ========================================================

    for li in ul.select("li"):

        text = norm_text(li)

        onclick = ""


        # ----------------------------------------------------
        # a 태그 우선
        # ----------------------------------------------------

        a = li.select_one(
            "a[onclick]"
        )

        if a:

            onclick = a.get(
                "onclick",
                ""
            )


        # ----------------------------------------------------
        # 혹시 다른 태그에 onclick이 있는 경우
        # ----------------------------------------------------

        if not onclick:

            for tag in li.find_all(
                attrs={
                    "onclick": True
                }
            ):

                temp = tag.get(
                    "onclick",
                    ""
                )

                if (
                    "fnUnvFileDownOne"
                    in temp
                ):

                    onclick = temp

                    break


        if not onclick:

            continue


        # ====================================================
        # fileId / fileSn 추출
        #
        # 작은따옴표 / 큰따옴표 모두 대응
        # ====================================================

        m = re.search(

            r"fnUnvFileDownOne"
            r"\(\s*"
            r"['\"]([^'\"]+)['\"]"
            r"\s*,\s*"
            r"['\"]([^'\"]+)['\"]",

            onclick

        )


        if not m:

            continue


        file_id = m.group(1)
        file_sn = m.group(2)


        # ====================================================
        # 파일 종류
        # ====================================================

        if (
            "대학입학전형" in text
            and
            "시행계획" in text
        ):

            plan_ids = (
                file_id,
                file_sn,
                text
            )


        elif (
            "수시" in text
            and
            "모집요강" in text
        ):

            susi_ids = (
                file_id,
                file_sn,
                text
            )


        elif (
            "정시" in text
            and
            "모집요강" in text
        ):

            jeongsi_ids = (
                file_id,
                file_sn,
                text
            )


    # ========================================================
    # 파일 다운로드
    # ========================================================

    file_buffers = {}


    target_files = [

        (
            "시행계획",
            plan_ids
        ),

        (
            "수시",
            susi_ids
        ),

        (
            "정시",
            jeongsi_ids
        )

    ]


    for label, ids in target_files:

        if not ids:

            continue


        f_id, f_sn, fname_text = ids


        params_file = {

            "fileId":
                f_id,

            "fileSn":
                f_sn,

            "menuId":
                MENU_ID,

            "downLogYn":
                "Y",

            "unvCd":
                str(unv_cd).zfill(7),

            "searchSyr":
                search_syr,

            "_":
                str(
                    int(
                        time.time() * 1000
                    )
                )

        }


        headers_file = {

            "User-Agent":
                "Mozilla/5.0",

            "Referer":
                (
                    f"{DETAIL_URL}"
                    f"?menuId={MENU_ID}"
                    f"&unvCd={unv_cd}"
                    f"&searchSyr={search_syr}"
                ),

            "X-Requested-With":
                "XMLHttpRequest",

            "Accept":
                "*/*"

        }


        try:

            r = session.get(

                DOWNLOAD_URL,

                params=params_file,

                headers=headers_file,

                timeout=60,

                verify=False

            )


            if r.status_code != 200:

                continue


            content = r.content


            if not content:

                continue


            content_type = (

                r.headers
                .get(
                    "Content-Type",
                    ""
                )
                .lower()

            )


            content_disp = (

                r.headers
                .get(
                    "Content-Disposition",
                    ""
                )

            )


            # =================================================
            # 실제 바이너리로 파일 형식 판정
            # =================================================

            if content.startswith(
                b"%PDF"
            ):

                ext = ".pdf"

                mime_type = (
                    "application/pdf"
                )


            elif (

                content.startswith(
                    b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
                )

                or

                b"HWP Document"
                in content[:512]

            ):

                ext = ".hwp"

                mime_type = (
                    "application/x-hwp"
                )


            elif content.startswith(
                b"PK"
            ):

                # HWPX
                if (
                    "hwpx" in content_disp.lower()
                    or
                    "hwpx" in content_type
                ):

                    ext = ".hwpx"

                # XLSX
                elif (
                    "xlsx" in content_disp.lower()
                    or
                    "spreadsheet" in content_type
                ):

                    ext = ".xlsx"

                else:

                    ext = ".zip"


                mime_type = (
                    "application/zip"
                )


            else:

                # Content-Disposition으로 보정
                lower_disp = (
                    content_disp.lower()
                )


                if ".pdf" in lower_disp:

                    ext = ".pdf"

                    mime_type = (
                        "application/pdf"
                    )


                elif ".hwpx" in lower_disp:

                    ext = ".hwpx"

                    mime_type = (
                        "application/octet-stream"
                    )


                elif ".hwp" in lower_disp:

                    ext = ".hwp"

                    mime_type = (
                        "application/octet-stream"
                    )


                elif ".xlsx" in lower_disp:

                    ext = ".xlsx"

                    mime_type = (
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )


                else:

                    ext = ".hwp"

                    mime_type = (
                        "application/octet-stream"
                    )


            fname = sanitize_filename(

                f"{univ_name}_"
                f"{label}_모집요강"
                f"{ext}"

            )


            file_buffers[label] = (

                content,

                fname,

                mime_type

            )


        except Exception as e:

            st.warning(
                f"{label} 파일 다운로드 실패: {e}"
            )


    return file_buffers


# ============================================================
# Streamlit 설정
# ============================================================

st.set_page_config(
    layout="wide"
)

st.title(
    "대학 입시자료 조회 및 다운로드"
)


# ============================================================
# 대학교 목록
# ============================================================

if not os.path.exists(
    UNIV_LIST_PATH
):

    st.error(
        f"{UNIV_LIST_PATH} 파일이 없습니다. "
        "깃허브에 포함시켜주세요."
    )

else:

    df = pd.read_excel(
        UNIV_LIST_PATH
    )


    if (
        "코드번호" not in df.columns
        or
        "학교명" not in df.columns
    ):

        st.error(
            "'코드번호'와 '학교명' 열이 필요합니다."
        )


    else:

        univ_list = (
            df["학교명"]
            .dropna()
            .tolist()
        )


        # ====================================================
        # 사이드바
        # ====================================================

        with st.sidebar:

            search_year = st.number_input(

                "학년도 입력",

                min_value=2000,

                max_value=2100,

                value=SEARCH_YEAR_DEFAULT,

                step=1

            )


            selected_univ = st.selectbox(

                "대학 선택",

                univ_list

            )


            types_options = (

                ["전체"]

                + list(
                    types_results.keys()
                )

                + list(
                    types_main.keys()
                )

            )


            selected_type = st.selectbox(

                "전형 선택",

                types_options

            )


        # ====================================================
        # 선택값 변경 시 이전 자료 삭제
        # ====================================================

        if (

            "selected_univ_prev"
            not in st.session_state

            or

            st.session_state.selected_univ_prev
            != selected_univ

            or

            st.session_state.get(
                "search_year_prev"
            )
            != search_year

            or

            st.session_state.get(
                "selected_type_prev"
            )
            != selected_type

        ):

            st.session_state.pop(
                "admission_data",
                None
            )

            st.session_state.pop(
                "file_buffers",
                None
            )

            st.session_state.selected_univ_prev = (
                selected_univ
            )

            st.session_state.search_year_prev = (
                search_year
            )

            st.session_state.selected_type_prev = (
                selected_type
            )


        # ====================================================
        # 화면 영역
        # ====================================================

        pdf_container = st.container()

        status_placeholder = st.empty()

        progress_bar = st.progress(0)


        # ====================================================
        # 크롤링 시작
        # ====================================================

        if st.button(
            "크롤링 시작"
        ):


            row = df[
                df["학교명"]
                == selected_univ
            ].iloc[0]


            unv_cd = str(
                row["코드번호"]
            ).strip()


            if unv_cd.endswith(
                ".0"
            ):

                unv_cd = (
                    unv_cd[:-2]
                )


            unv_cd = re.sub(
                r"\D",
                "",
                unv_cd
            )


            unv_cd = unv_cd.zfill(
                7
            )


            st.session_state.admission_data = {}

            st.session_state.file_buffers = {}


            # =================================================
            # 선택된 전형 결정
            # =================================================

            all_types = {}


            if selected_type == "전체":

                all_types = {

                    **types_main,

                    **types_results

                }


            elif selected_type in types_main:

                all_types = {

                    selected_type:
                        types_main[
                            selected_type
                        ]

                }


            elif selected_type in types_results:

                all_types = {

                    selected_type:
                        types_results[
                            selected_type
                        ]

                }


            total = len(
                all_types
            )


            # =================================================
            # 주요사항 / 입시결과
            # =================================================

            for i, (
                name,
                codes
            ) in enumerate(
                all_types.items(),
                1
            ):


                status_placeholder.info(

                    f"{name} 크롤링 중..."
                    f" ({i}/{total})"

                )


                # 주요사항
                if name.endswith(
                    "(주요사항)"
                ):

                    data_chunk = (
                        crawl_major_items(

                            unv_cd,

                            search_year,

                            name,

                            codes

                        )
                    )


                # 입시결과
                else:

                    data_chunk = (
                        crawl_admission_results_chunk(

                            unv_cd,

                            search_year,

                            name,

                            codes

                        )
                    )


                st.session_state.admission_data.update(
                    data_chunk
                )


                if total:

                    progress_bar.progress(
                        i / total
                    )


            # =================================================
            # 모집요강
            # =================================================

            status_placeholder.info(

                f"{search_year}학년도 "
                "모집요강 파일 확인 중..."

            )


            st.session_state.file_buffers = (

                extract_and_download_files(

                    unv_cd,

                    search_year,

                    selected_univ

                )

            )


            progress_bar.progress(
                1.0
            )


            status_placeholder.success(
                "크롤링 완료! ✅"
            )


        # ====================================================
        # 결과 표시
        # ====================================================

        if (

            "admission_data"
            in st.session_state

            and

            st.session_state.admission_data

        ):


            type_order = [

                (
                    "학생부종합",
                    "2️⃣ 학생부종합전형"
                ),

                (
                    "학생부교과",
                    "3️⃣ 학생부교과전형"
                ),

                (
                    "수능",
                    "4️⃣ 수능위주전형"
                )

            ]


            for (
                type_name,
                header_name
            ) in type_order:


                if (

                    type_name
                    in st.session_state.admission_data

                    or

                    f"{type_name}(주요사항)"
                    in st.session_state.admission_data

                ):


                    st.markdown(
                        f"## {header_name}"
                    )


                    # ----------------------------------------
                    # 주요사항
                    # ----------------------------------------

                    main_name = (
                        f"{type_name}(주요사항)"
                    )


                    if (
                        main_name
                        in st.session_state.admission_data
                    ):


                        st.markdown(

                            f"### 📌 "
                            f"{search_year}학년도 "
                            f"전형별 주요사항"

                        )


                        df_main = (

                            st.session_state
                            .admission_data[
                                main_name
                            ]

                        )


                        st.dataframe(

                            wrap_long_text(
                                df_main,
                                max_len=50
                            ),

                            use_container_width=True

                        )


                    # ----------------------------------------
                    # 입시결과
                    # ----------------------------------------

                    result_name = type_name


                    if (
                        result_name
                        in st.session_state.admission_data
                    ):


                        st.markdown(

                            f"### 📊 "
                            f"{search_year - 1}"
                            f"학년도 전형 결과"

                        )


                        df_result = (

                            st.session_state
                            .admission_data[
                                result_name
                            ]

                        )


                        st.dataframe(

                            wrap_long_text(
                                df_result,
                                max_len=50
                            ),

                            use_container_width=True

                        )


            # =================================================
            # Excel 다운로드
            # =================================================

            excel_buffer = BytesIO()


            with pd.ExcelWriter(

                excel_buffer,

                engine="openpyxl"

            ):


                for (
                    type_name,
                    _
                ) in type_order:


                    main_name = (
                        f"{type_name}(주요사항)"
                    )


                    if (
                        main_name
                        in st.session_state.admission_data
                    ):


                        st.session_state.admission_data[
                            main_name
                        ].to_excel(

                            writer,

                            sheet_name=
                                sanitize_filename(
                                    main_name
                                )[:31],

                            index=False,

                            header=False

                        )


                    result_name = type_name


                    if (
                        result_name
                        in st.session_state.admission_data
                    ):


                        st.session_state.admission_data[
                            result_name
                        ].to_excel(

                            writer,

                            sheet_name=
                                sanitize_filename(
                                    result_name
                                )[:31],

                            index=False,

                            header=False

                        )


            excel_buffer.seek(0)


            st.download_button(

                label=
                    "📥 입시결과 다운로드",

                data=
                    excel_buffer,

                file_name=

                    (
                        f"{sanitize_filename(selected_univ)}_"
                        f"{search_year - 1}년_"
                        f"대학입시결과.xlsx"
                    ),

                mime=
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            )


            # =================================================
            # 모집요강 다운로드
            # =================================================

            with pdf_container:


                if (
                    st.session_state.file_buffers
                ):


                    st.markdown(
                        "### 1️⃣ 모집요강 다운로드"
                    )


                    for (
                        label,
                        (
                            content,
                            fname,
                            mime_type
                        )
                    ) in (

                        st.session_state
                        .file_buffers
                        .items()

                    ):


                        st.download_button(

                            label=
                                f"📄 {label} 다운로드",

                            data=
                                content,

                            file_name=
                                fname,

                            mime=
                                mime_type

                        )


                else:

                    st.warning(

                        f"{search_year}학년도 "
                        "모집요강 파일이 없습니다."

                    )
