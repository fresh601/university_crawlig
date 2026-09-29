import os, re, time, unicodedata, warnings
from io import BytesIO
import requests
import pandas as pd
from bs4 import BeautifulSoup
import streamlit as st

warnings.filterwarnings('ignore', message='Unverified HTTPS request')

BASE='https://www.adiga.kr'
POPUP_URL=f'{BASE}/uct/acd/ade/criteriaAndResultPopup.do'
RESULT_URL=f'{BASE}/uct/acd/ade/criteriaAndResultItemNewAjax.do'
MAJOR_URL=f'{BASE}/uct/acd/ade/criteriaAndResultItemAjax.do'
DETAIL_URL=f'{BASE}/ucp/uvt/uni/univDetail.do'
DOWNLOAD_URL=f'{BASE}/cmm/com/file/fileDown.do'
MENU_ID='PCUVTINF2000'
UNIV_LIST_PATH='대학교별 코드.xlsx'
SEARCH_YEAR_DEFAULT=2027

RESULT_TYPES={
    '학생부종합': '20',
    '학생부교과': '30',
    '수능': '40',
}
MAJOR_TYPES={
    '학생부종합(주요사항)': {'upcd':'20','artclcd':'21'},
    '학생부교과(주요사항)': {'upcd':'30','artclcd':'31'},
    '수능(주요사항)': {'upcd':'40','artclcd':'41'},
}

def safe_name(s):
    s=unicodedata.normalize('NFKC',str(s))
    return re.sub(r'\s+',' ',re.sub(r'[<>:"/\\|?*\x00-\x1F]',' ',s)).strip()

def clean(s):
    return ' '.join(str(s).replace('\xa0',' ').split())

def session_new():
    s=requests.Session()
    s.headers.update({'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/153.0.0.0 Safari/537.36','Accept-Language':'ko-KR,ko;q=0.9,en-US;q=0.8'})
    try: s.get(BASE,timeout=20,verify=False)
    except Exception: pass
    return s

def university_page(s, unv, year):
    r=s.get(POPUP_URL,params={'searchSyr':year,'unvCd':unv,'tsrdCmphSlcnArtclUpCd':'20'},timeout=40,verify=False)
    r.raise_for_status()
    soup=BeautifulSoup(r.text,'html.parser')
    tag=soup.select_one('input[name="_csrf"]')
    return r.text, (tag.get('value','') if tag else ''), r.url

def parse_table(table):
    rows=table.find_all('tr')
    if not rows: return None
    occupied={}; max_col=0
    for ri,row in enumerate(rows):
        cells=row.find_all(['th','td'],recursive=False) or row.find_all(['th','td'])
        ci=0
        for cell in cells:
            while (ri,ci) in occupied: ci+=1
            text=clean(cell.get_text(' ',strip=True))
            try: rs=max(1,int(cell.get('rowspan',1)))
            except: rs=1
            try: cs=max(1,int(cell.get('colspan',1)))
            except: cs=1
            for r in range(ri,ri+rs):
                for c in range(ci,ci+cs):
                    occupied.setdefault((r,c),text)
            ci+=cs; max_col=max(max_col,ci)
    matrix=[[occupied.get((r,c),'') for c in range(max_col)] for r in range(len(rows))]
    return pd.DataFrame(matrix)

def tables_from_html(html):
    soup=BeautifulSoup(html,'html.parser')
    out=[]
    for t in soup.find_all('table'):
        try:
            d=parse_table(t)
            if d is not None and not d.empty: out.append(d)
        except Exception: pass
    return out

def fetch_result(s,unv,year,upcd,csrf,referer):
    data={'searchSyr':str(year),'unvCd':unv,'tsrdCmphSlcnArtclUpCd':upcd,'compUnvCd':''}
    headers={'Accept':'application/json, text/plain, */*','Content-Type':'application/x-www-form-urlencoded; charset=UTF-8','Origin':BASE,'Referer':referer,'X-Requested-With':'XMLHttpRequest'}
    if csrf: headers['X-CSRF-TOKEN']=csrf
    r=s.post(RESULT_URL,data=data,headers=headers,timeout=45,verify=False)
    r.raise_for_status(); return r.text

def fetch_major(s,unv,year,codes,csrf,referer):
    data={'_csrf':csrf,'searchSyr':str(year),'searchStdClsfRgnCn':'','searchUnvNm':'','unvCd':unv,'compUnvCd':'','searchUnvComp':'0','tsrdCmphSlcnArtclUpCd':codes['upcd'],'tsrdCmphSlcnArtclCd':codes['artclcd']}
    headers={'X-CSRF-TOKEN':csrf,'X-Requested-With':'XMLHttpRequest','Referer':referer}
    r=s.post(MAJOR_URL,data=data,headers=headers,timeout=45,verify=False)
    r.raise_for_status(); return r.text

def find_files(html):
    soup=BeautifulSoup(html,'html.parser'); ul=soup.select_one('ul#fileResult')
    found={'시행계획':None,'수시':None,'정시':None}
    if not ul:return found
    for li in ul.select('li'):
        a=li.select_one('a[onclick]'); span=li.select_one('span')
        if not a: continue
        text=clean(span.get_text(' ',strip=True) if span else li.get_text(' ',strip=True)); oc=a.get('onclick','')
        m=re.search(r"fnUnvFileDownOne\(\s*['\"]([^'\"]+)['\"]\s*,\s*['\"]([^'\"]+)['\"]",oc)
        if not m: continue
        pair=(m.group(1),m.group(2),text)
        if '대학입학전형' in text and '시행계획' in text: found['시행계획']=pair
        elif '수시' in text and '모집요강' in text: found['수시']=pair
        elif '정시' in text and '모집요강' in text: found['정시']=pair
    return found

def file_type(content,headers):
    if content.startswith(b'%PDF'): return '.pdf','application/pdf'
    if content.startswith(b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'): return '.hwp','application/x-hwp'
    if content.startswith(b'PK'):
        cd=str(headers.get('Content-Disposition','')).lower()
        if '.hwpx' in cd:return '.hwpx','application/hwp+zip'
        if '.xlsx' in cd:return '.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        try:
            import zipfile
            with zipfile.ZipFile(BytesIO(content)) as z:
                names=z.namelist()
                if any(x.startswith('Contents/') for x in names): return '.hwpx','application/hwp+zip'
                if any(x.startswith('xl/') for x in names): return '.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        except Exception: pass
        return '.zip','application/zip'
    return '.bin','application/octet-stream'

def download_recruitment(s,unv,year,univ,detail_html):
    found=find_files(detail_html); result={}
    for label,item in found.items():
        if not item: continue
        fid,fsn,source=item
        params={'fileId':fid,'fileSn':fsn,'menuId':MENU_ID,'downLogYn':'Y','unvCd':unv,'searchSyr':str(year),'_':str(int(time.time()*1000))}
        headers={'User-Agent':'Mozilla/5.0','Referer':f'{DETAIL_URL}?menuId={MENU_ID}&unvCd={unv}&searchSyr={year}','X-Requested-With':'XMLHttpRequest'}
        try:
            r=s.get(DOWNLOAD_URL,params=params,headers=headers,timeout=90,verify=False); r.raise_for_status()
            ext,mime=file_type(r.content,r.headers)
            result[label]=(r.content,safe_name(f'{univ}_{label}_모집요강{ext}'),mime)
        except Exception as e: st.warning(f'{label} 다운로드 실패: {e}')
    return result

def wrap_df(df,n=50):
    d=df.copy()
    for c in d.columns: d[c]=d[c].map(lambda x:'\n'.join(str(x)[i:i+n] for i in range(0,len(str(x)),n)))
    return d

st.set_page_config(page_title='대학 입시자료 조회',page_icon='🎓',layout='wide')
st.title('🎓 대학 입시자료 조회 및 다운로드')

if not os.path.exists(UNIV_LIST_PATH): st.error(f'{UNIV_LIST_PATH} 파일이 없습니다. GitHub에 포함시켜주세요.'); st.stop()
df=pd.read_excel(UNIV_LIST_PATH)
if '코드번호' not in df.columns or '학교명' not in df.columns: st.error("'코드번호'와 '학교명' 열이 필요합니다."); st.stop()
df=df.dropna(subset=['학교명','코드번호']).copy(); df['학교명']=df['학교명'].astype(str)

with st.sidebar:
    st.header('🔎 조회 조건')
    year=st.number_input('학년도',2000,2100,SEARCH_YEAR_DEFAULT,1)
    univ=st.selectbox('대학 선택',df['학교명'].tolist())
    typ=st.selectbox('전형 선택',['전체']+list(RESULT_TYPES.keys()))
    st.info(f'주요사항: {year}학년도\n\n입시결과: {year-1}학년도\n\n모집요강: {year}학년도')

key=(univ,int(year),typ)
if st.session_state.get('key')!=key:
    for k in ['data','files','error']: st.session_state.pop(k,None)
    st.session_state.key=key

if st.button('🚀 크롤링 시작',type='primary',use_container_width=True):
    row=df[df['학교명']==univ].iloc[0]; unv=str(row['코드번호']).split('.')[0].zfill(7)
    st.session_state.data={}; st.session_state.files={}
    s=session_new(); status=st.empty(); bar=st.progress(0)
    try:
        detail,csrf,referer=university_page(s,unv,int(year))
    except Exception as e:
        st.error(f'대학 페이지 조회 실패: {e}'); st.stop()

    if typ=='전체':
        jobs=[(f'{n}(주요사항)',MAJOR_TYPES[f'{n}(주요사항)']) for n in RESULT_TYPES]
        jobs += [(n,{'upcd':u}) for n,u in RESULT_TYPES.items()]
    else:
        jobs=[(f'{typ}(주요사항)',MAJOR_TYPES[f'{typ}(주요사항)']),(typ,{'upcd':RESULT_TYPES[typ]})]

    for i,(name,codes) in enumerate(jobs,1):
        status.info(f'{name} 조회 중... ({i}/{len(jobs)})')
        try:
            if name.endswith('(주요사항)'): html=fetch_major(s,unv,int(year),codes,csrf,referer)
            else: html=fetch_result(s,unv,int(year),codes['upcd'],csrf,referer)
            tabs=tables_from_html(html)
            if tabs:
                frames=[]
                for t in tabs:
                    frames.append(t); frames.append(pd.DataFrame([['']*t.shape[1]]))
                st.session_state.data[name]=pd.concat(frames,ignore_index=True)
        except Exception as e: st.warning(f'{name} 실패: {e}')
        bar.progress(i/len(jobs))

    status.info('모집요강 다운로드 중...')
    st.session_state.files=download_recruitment(s,unv,int(year),univ,detail)
    status.success('✅ 크롤링 완료')

if st.session_state.get('data'):
    order=[('학생부종합','2️⃣ 학생부종합전형'),('학생부교과','3️⃣ 학생부교과전형'),('수능','4️⃣ 수능위주전형')]
    for n,title in order:
        main=f'{n}(주요사항)'
        if main in st.session_state.data or n in st.session_state.data:
            st.markdown(f'## {title}')
            if main in st.session_state.data:
                st.markdown(f'### 📌 {year}학년도 전형별 주요사항')
                st.dataframe(wrap_df(st.session_state.data[main]),use_container_width=True,height=500,hide_index=True)
            if n in st.session_state.data:
                st.markdown(f'### 📊 {year-1}학년도 전형 결과')
                st.dataframe(wrap_df(st.session_state.data[n]),use_container_width=True,height=500,hide_index=True)

    buf=BytesIO()
    with pd.ExcelWriter(buf,engine='openpyxl') as w:
        for n,_ in order:
            for sheet in [f'{n}(주요사항)',n]:
                if sheet in st.session_state.data: st.session_state.data[sheet].to_excel(w,sheet_name=safe_name(sheet)[:31],index=False,header=False)
    buf.seek(0)
    st.download_button('📊 입시결과 + 주요사항 Excel 다운로드',buf.getvalue(),f'{safe_name(univ)}_{year}학년도_입시자료.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)

if 'files' in st.session_state:
    st.markdown(f'### 📚 {year}학년도 모집요강')
    if st.session_state.files:
        cols=st.columns(len(st.session_state.files))
        for col,(label,(content,fname,mime)) in zip(cols,st.session_state.files.items()):
            with col: st.download_button(f'📄 {label} 다운로드',content,fname,mime,use_container_width=True); st.caption(fname)
    else: st.warning('모집요강 파일을 찾지 못했습니다.')

with st.expander('ℹ️ 조회 기준'):
    st.markdown(f'- 조회 학년도: {year}학년도\n- 주요사항: {year}학년도\n- 입시결과: {year-1}학년도\n- 모집요강: {year}학년도\n- 자료 출처: 대입정보포털 어디가')
