import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
from datetime import datetime
from io import BytesIO
from openpyxl.styles import Font
import openpyxl
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials

# QR/바코드 스캔 라이브러리 (파이썬 백엔드 분석용)
try:
    from PIL import Image
    import cv2
    from pyzbar.pyzbar import decode
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

# [설정] 작업자 명단
worker_list = ["박경섭", "무고사", "재르소", "김동헌"] 

st.set_page_config(
    page_title="VISION DATA KEY-IN SYSTEM ----- (by. Romero)", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------
# 마법 코드 1: UI 숨김 및 태블릿 앱 최적화
# ----------------------------------------------------
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
body { overscroll-behavior-y: none !important; }
* {
    -webkit-user-select: none; -ms-user-select: none; user-select: none; 
    -webkit-tap-highlight-color: transparent !important; 
}
input, textarea, select {
    -webkit-user-select: auto !important; -ms-user-select: auto !important; user-select: auto !important;
}
::-webkit-scrollbar { display: none; }
.block-container {
    padding-top: 1rem !important; padding-bottom: 1rem !important;
    padding-left: 1.5rem !important; padding-right: 1.5rem !important;
}
button[kind="primary"] {
    background-color: #4b6584 !important; color: white !important; border: none !important;
    font-size: 16px !important; font-weight: bold !important; padding: 10px !important;
}
button[kind="primary"]:hover { background-color: #3b5068 !important; }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

components.html(
    """
    <script>
    if (window.parent && !window.parent.appPluginLoaded) {
        window.parent.appPluginLoaded = true;
        const doc = window.parent.document;
        const head = doc.head;
        const metaTags = [
            { name: "mobile-web-app-capable", content: "yes" },
            { name: "apple-mobile-web-app-capable", content: "yes" },
            { name: "apple-mobile-web-app-status-bar-style", content: "black-translucent" },
            { name: "viewport", content: "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no" }
        ];
        metaTags.forEach(tag => {
            let meta = doc.createElement('meta');
            meta.name = tag.name;
            meta.content = tag.content;
            head.appendChild(meta);
        });

        const disableKeyboard = () => {
            if (!doc) return;
            doc.querySelectorAll('div[data-baseweb="select"] input').forEach(el => {
                if (el.getAttribute('inputmode') !== 'none') el.setAttribute('inputmode', 'none');
            });
            doc.querySelectorAll('input').forEach(el => {
                const placeholder = el.getAttribute('placeholder') || '';
                const hasPopup = el.hasAttribute('aria-haspopup');
                if (placeholder.includes('YYYY') || placeholder.includes('HH:MM') || hasPopup) {
                    if (el.getAttribute('inputmode') !== 'none') el.setAttribute('inputmode', 'none');
                }
            });
        };
        const observer = new MutationObserver(() => { disableKeyboard(); });
        if (window.parent.document.body) {
            observer.observe(window.parent.document.body, { childList: true, subtree: true });
        }
        disableKeyboard();
    }
    </script>
    """, height=0, width=0
)

# ----------------------------------------------------
# 터치형 박스(Grid) 렌더링 함수
# ----------------------------------------------------
def render_grid_buttons(options, state_key, columns):
    if state_key not in st.session_state:
        valid_opts = [o for o in options if o.strip()]
        st.session_state[state_key] = valid_opts[0] if valid_opts else ""
        
    rows = [options[i:i+columns] for i in range(0, len(options), columns)]
    for row_opts in rows:
        cols = st.columns(columns)
        for i, opt in enumerate(row_opts):
            with cols[i]:
                if opt.strip() == "":
                    st.write("") 
                else:
                    btn_type = "primary" if st.session_state[state_key] == opt else "secondary"
                    if st.button(opt, key=f"btn_{state_key}_{opt}", type=btn_type, use_container_width=True):
                        st.session_state[state_key] = opt
                        st.rerun()

# ----------------------------------------------------
# 구글 스프레드시트 연동
# ----------------------------------------------------
SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

EXCEL_COLUMNS = [
    "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", 
    "모델명(MI)", "도금구분", "UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", 
    "불량수량", "양품률", "양품율(전/배 포함)", "완전불량률", "전면불량률", "배면불량률", 
    "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타", "OQC", "비고", 
    "도장라인", "도장일", "도장순서", "입고일", "LOT NO.", "CLIP", "BASE", "COVER", 
    "조립기", "월", "작업자"
]

SPREADSHEET_ID = "1DeMJJkuq7bYa4XNK_NbkqZ-vOJKqGhmYXIvHm3yJl8E"
TAB_NAME = "VISION_DATA_DB"

if "google_credentials" not in st.secrets:
    st.error("구글 스프레드시트 보안 키(Secrets)가 설정되지 않았습니다!")
    st.stop()

@st.cache_resource
def get_sheet():
    try:
        creds_data = st.secrets["google_credentials"]
        clean_data = creds_data.strip().strip("'").strip('"') if isinstance(creds_data, str) else dict(creds_data)
        creds_dict = json.loads(clean_data, strict=False) if isinstance(creds_data, str) else clean_data
        if "private_key" in creds_dict: creds_dict["private_key"] = creds_dict["private_key"].replace('\\n', '\n')
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
        
        doc = gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
        try:
            return doc.worksheet(TAB_NAME)
        except gspread.exceptions.WorksheetNotFound:
            return doc.sheet1
            
    except Exception as e:
        st.error(f"구글 연결 초기화 에러 (권한이 없거나 키가 잘못되었습니다): {e}")
        return None

@st.cache_data(ttl=60)
def load_data():
    sheet = get_sheet()
    if sheet is None: return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        try: 
            raw_data = sheet.get_all_values()
        except Exception as e: 
            st.error(f"구글 시트 데이터 읽기 실패: {e}")
            return pd.DataFrame(columns=EXCEL_COLUMNS)
            
        valid_data = [row for row in raw_data if any(str(cell).strip() for cell in row)]
            
        if not valid_data or len(valid_data) < 2: 
            return pd.DataFrame(columns=EXCEL_COLUMNS)
            
        header_idx = 0
        for i, row in enumerate(valid_data[:10]):
            row_str = "".join(str(c).replace(" ", "") for c in row)
            if "날짜" in row_str or "교대" in row_str or "모델명" in row_str:
                header_idx = i
                break
                
        headers = [str(h).strip() for h in valid_data[header_idx]]
        
        df = pd.DataFrame(valid_data[header_idx+1:])
        if df.empty:
            return pd.DataFrame(columns=EXCEL_COLUMNS)
            
        if len(df.columns) > len(headers):
            df = df.iloc[:, :len(headers)]
        df.columns = headers[:len(df.columns)]
        
        clean_headers = {c.replace(" ", "").upper(): c for c in df.columns}
        
        result_df = pd.DataFrame(index=df.index)
        for col in EXCEL_COLUMNS:
            col_key = col.replace(" ", "").upper()
            if col_key in clean_headers:
                actual_col = clean_headers[col_key]
                col_data = df[actual_col]
                result_df[col] = col_data.iloc[:, 0] if isinstance(col_data, pd.DataFrame) else col_data
            else:
                result_df[col] = "" 
                
        return result_df
    except Exception as e:
        st.error(f"데이터 변환 처리 중 에러 발생: {e}")
        return pd.DataFrame(columns=EXCEL_COLUMNS)

def clean_for_gsheet(df):
    df_clean = df.copy()
    for col in df_clean.columns:
        df_clean[col] = df_clean[col].apply(lambda x: "" if str(x).strip().lower() in ["nan", "nat", "none", "<na>", "inf", "-inf"] else str(x))
    return df_clean

def save_data_append(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        try: header_check = sheet.row_values(1)
        except Exception: header_check = []
        if not header_check: sheet.append_row(EXCEL_COLUMNS, value_input_option='USER_ENTERED')
        
        records_to_insert = []
        for _, row in df.iterrows():
            row_data = ["" if str(row.get(col, "")).strip().lower() in ["nan", "nat", "none", "inf", "-inf"] else str(row.get(col, "")).strip() for col in EXCEL_COLUMNS]
            records_to_insert.append(row_data)

        try: sheet.append_rows(records_to_insert, value_input_option='USER_ENTERED')
        except Exception:
            for r_data in records_to_insert: sheet.append_row(r_data, value_input_option='USER_ENTERED')
        load_data.clear() 
        return True
    except Exception as e:
        st.error(f"데이터 저장 오류: {e}")
        return False

def save_data_overwrite(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        try: sheet.clear()
        except Exception: pass
        records_to_insert = [EXCEL_COLUMNS] 
        for _, row in df.iterrows():
            row_data = ["" if str(row.get(col, "")).strip().lower() in ["nan", "nat", "none", "inf", "-inf"] else str(row.get(col, "")).strip() for col in EXCEL_COLUMNS]
            records_to_insert.append(row_data)
        try: sheet.update(values=records_to_insert, range_name='A1', value_input_option='USER_ENTERED')
        except TypeError: sheet.update('A1', records_to_insert, value_input_option='USER_ENTERED')
        load_data.clear()
        return True
    except Exception as e:
        st.error(f"데이터 덮어쓰기 오류: {e}")
        return False

# ----------------------------------------------------
# 팝업(모달) & Python 이미지 분석 기반 QR 스캐너
# ----------------------------------------------------
@st.dialog("SBL Warning!")
def show_sbl_warning(defect_type, rate):
    st.markdown(f"### [{defect_type}] 불량 제품을 별도 보관 조치 하세요.")
    st.error(f"현재 1차검사 공정의 {defect_type}율이 **{rate:.1f}%** 로 기준치(5.0%) 초과하였습니다.")
    if st.button("확인 완료 (닫기)", key=f"btn_close_{defect_type}"):
        st.rerun()

@st.dialog("관리자 인증")
def show_password_dialog():
    st.markdown("분석 데이터를 확인하려면 관리자 비밀번호를 입력하세요.")
    pwd = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
    if st.button("확인", type="primary", use_container_width=True):
        if pwd == "6233":
            st.session_state.current_page = "analysis"
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")

@st.dialog("카메라 촬영 및 자동 분석")
def open_camera_qr_scanner():
    if not QR_AVAILABLE:
        st.error("QR 라이브러리(opencv-python-headless, pyzbar)가 설치되지 않았습니다.")
        return
        
    st.info("팁: QR 코드가 화면에 선명하게 보일 때 사진을 찍어주세요.")
    
    tab1, tab2 = st.tabs(["카메라 촬영", "갤러리 앨범"])
    
    target_img = None
    with tab1:
        img_buffer = st.camera_input("카메라 촬영")
        if img_buffer: target_img = img_buffer
        
    with tab2:
        uploaded_img = st.file_uploader("앨범에서 사진 선택", type=['png', 'jpg', 'jpeg'])
        if uploaded_img: target_img = uploaded_img
        
    if target_img:
        with st.spinner("이미지 분석 중..."):
            try:
                image = Image.open(target_img)
                cv2_img = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
                
                gray = cv2.cvtColor(cv2_img, cv2.COLOR_BGR2GRAY)
                clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8,8))
                enhanced_gray = clahe.apply(gray)
                
                objs = decode(cv2_img)
                if not objs:
                    objs = decode(enhanced_gray)
                    
                if objs:
                    raw_data = objs[0].data.decode('utf-8')
                    
                    parts = [p for p in raw_data.split('$') if p]
                    lot_val = parts[-1] if parts else raw_data
                    parsed_date = None
                    
                    if len(parts) >= 2:
                        date_str_candidate = parts[-2]
                        date_str = date_str_candidate[-8:]
                        if date_str.isdigit():
                            try:
                                parsed_date = datetime.strptime(date_str, "%Y%m%d").date()
                            except ValueError:
                                pass
                    
                    if parsed_date:
                        st.success(f"인식 성공! LOT: {lot_val} / 입고일: {parsed_date.strftime('%Y/%m/%d')}")
                    else:
                        st.success(f"인식 성공! LOT: {lot_val} (날짜 미인식)")
                        
                    st.caption(f"원본 바코드: {raw_data}")
                    
                    if st.button("입력창에 적용 및 닫기", type="primary", use_container_width=True):
                        st.session_state.temp_lot = lot_val
                        if parsed_date:
                            st.session_state.temp_date = parsed_date
                        st.rerun()
                else:
                    st.error("QR 코드를 찾을 수 없습니다. 초점을 맞춰서 다시 촬영해주세요.")
            except Exception as e:
                st.error(f"분석 중 오류 발생: {e}")

# ==========================================
# 화면 전환 및 상태 관리
# ==========================================
if "current_page" not in st.session_state: st.session_state.current_page = "input"
if "lot_input_field" not in st.session_state: st.session_state.lot_input_field = ""
if "in_date_field" not in st.session_state: st.session_state.in_date_field = datetime.now().date()

# ==========================================
# 종합 분석 데이터 화면 
# ==========================================
def render_analysis_page():
    st.markdown("""
        <div style='background: linear-gradient(135deg, #0f172a 0%, #020617 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1e293b;'>
            <h2 style='color: #f8fafc; margin: 0; font-weight: 600;'>종합 생산 데이터 분석</h2>
        </div>
    """, unsafe_allow_html=True)

    col_btn1, col_btn2 = st.columns([0.8, 0.2])
    with col_btn1:
        if st.button("뒤로 가기 (데이터 입력 화면으로)", type="primary"):
            st.session_state.current_page = "input"
            st.rerun()
    with col_btn2:
        if st.button("🔄 최신 데이터 새로고침", use_container_width=True):
            load_data.clear()
            st.rerun()
            
    active_sheet = get_sheet()
    if active_sheet:
        st.caption(f"🔗 데이터 소스: **[{active_sheet.title}]** 탭 연동 완료")
            
    df = load_data().copy()
    if df.empty:
        st.warning("분석할 저장된 데이터가 없습니다. 우측 상단의 '최신 데이터 새로고침' 버튼을 눌러주세요.")
        return

    num_cols = ["검사 수량", "양품수량", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

    current_year = datetime.now().year
    def parse_custom_date(val):
        try:
            val_str = str(val).strip()
            if '/' in val_str and len(val_str) <= 5: 
                return pd.to_datetime(f"{current_year}/{val_str}", format="%Y/%m/%d")
            return pd.to_datetime(val_str) 
        except:
            return pd.NaT

    if '날짜' in df.columns:
        df['분석용_날짜'] = df['날짜'].apply(parse_custom_date)
    else:
        df['분석용_날짜'] = pd.NaT

    if '시작시간' in df.columns: df['시간대'] = df['시작시간'].str[:13] + "시"
    else: df['시간대'] = df['날짜']

    f_head_col1, f_head_col2 = st.columns([0.95, 0.05])
    with f_head_col1: st.markdown("##### 상세 분석 필터")
    with f_head_col2:
        with st.popover("🎨"):
            pc1, pc2, pc3, pc4, pc5 = st.columns(5)
            with pc1: c_yield1 = st.color_picker("양품", "#002b5e")
            with pc2: c_comp = st.color_picker("완전불량", "#b22222")
            with pc3: c_front = st.color_picker("전면불량", "#ed7d31")
            with pc4: c_rear = st.color_picker("배면불량", "#00b050")
            with pc5: c_offset = st.color_picker("옵셋불량", "#7030a0")

    available_dates = df.dropna(subset=['분석용_날짜']).sort_values('분석용_날짜', ascending=False)['날짜'].unique()
    
    col_opt1, col_opt2 = st.columns(2)
    with col_opt1: date_filter_mode = st.radio("**분석 기간 설정**", ["전체 누적 데이터", "단일 일자 선택", "특정 기간 지정 검색"], horizontal=True)
    with col_opt2: x_axis_mode = st.radio("**분석 기준 (X축)**", ["일별 (날짜 기준)", "시간별 (시작시간 기준)"], horizontal=True)
    
    if date_filter_mode == "단일 일자 선택":
        selected_date = st.selectbox("**분석할 근무일자를 선택하세요**", available_dates)
        if selected_date: df = df[df['날짜'] == selected_date]
    elif date_filter_mode == "특정 기간 지정 검색":
        d_col1, d_col2 = st.columns(2)
        try: 
            min_date = df['분석용_날짜'].min().date()
            max_date = df['분석용_날짜'].max().date()
        except: 
            min_date = max_date = datetime.now().date()
            
        with d_col1: start_date_filter = st.date_input("**시작 일자**", value=min_date)
        with d_col2: end_date_filter = st.date_input("**종료 일자**", value=max_date)
        
        start_dt = pd.to_datetime(start_date_filter)
        end_dt = pd.to_datetime(end_date_filter)
        df = df[(df['분석용_날짜'] >= start_dt) & (df['분석용_날짜'] <= end_dt)]

    f_col1, f_col2, f_col3, f_col4 = st.columns(4)
    with f_col1: selected_model = st.selectbox("**모델명**", ["전체"] + sorted(df['모델명(MI)'].dropna().astype(str).unique()))
    with f_col2: selected_unit = st.selectbox("**호기**", ["전체"] + sorted(df['호기'].dropna().astype(str).unique()))
    with f_col3: selected_category = st.selectbox("**검사구분**", ["전체"] + sorted(df['구분'].dropna().astype(str).unique()))
    with f_col4: selected_shift = st.selectbox("**교대**", ["전체"] + sorted(df['교대'].dropna().astype(str).unique()))

    if selected_model != "전체": df = df[df['모델명(MI)'] == selected_model]
    if selected_unit != "전체": df = df[df['호기'] == selected_unit]
    if selected_category != "전체": df = df[df['구분'] == selected_category]
    if selected_shift != "전체": df = df[df['교대'] == selected_shift]

    if df.empty or df["검사 수량"].sum() == 0:
        st.info("조건에 맞는 데이터가 없습니다.")
        return

    base_col = '시간대' if "시간별" in x_axis_mode else '날짜'
    group_cols = [base_col]
    if selected_model == "전체": group_cols.append('모델명(MI)')

    df_grouped = df.groupby(group_cols)[num_cols].sum().reset_index().sort_values(base_col)
    df_grouped['양품률'] = (df_grouped['양품수량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['양품율(전/배 포함)'] = ((df_grouped['양품수량'] + df_grouped['전면불량'] + df_grouped['배면불량']) / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['불량율'] = (df_grouped['불량수량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['완전불량률'] = (df_grouped['완전불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['전면불량률'] = (df_grouped['전면불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['배면불량률'] = (df_grouped['배면불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)
    df_grouped['옵셋불량율'] = (df_grouped['옵셋불량'] / df_grouped['검사 수량'] * 100).fillna(0).round(1)

    is_single_x = len(df_grouped[base_col].unique()) == 1

    # 💡 [핵심 디자인 엔진] 참조 파일 스타일 완벽 이식 (가시성 및 텍스트 섀도우)
    def create_single_chart(title, metric, base_color):
        fig = go.Figure()
        
        if "양품" in title or "수율" in title:
            fc, sc = ("black", "#e0e0e0")
        else:
            fc, sc = ("white", "#000000")
            
        ts = f"color:{fc}; text-shadow: -1px -1px 0 {sc}, 1px -1px 0 {sc}, -1px 1px 0 {sc}, 1px 1px 0 {sc};"

        def get_tv(x):
            if pd.isna(x) or x == "": return ""
            val = float(x)
            if abs(val) < 0.05: return ""
            return f"<b><span style='{ts}'>{val:.1f}%</span></b>"

        if selected_model == "전체":
            colors = px.colors.qualitative.Plotly 
            for i, model in enumerate(df_grouped['모델명(MI)'].unique()):
                m_data = df_grouped[df_grouped['모델명(MI)'] == model]
                tv = m_data[metric].apply(get_tv)
                if is_single_x: 
                    fig.add_trace(go.Bar(name=str(model), x=m_data[base_col], y=m_data[metric], marker_color=colors[i%len(colors)], text=tv, textposition='inside', insidetextanchor='middle'))
                else: 
                    fig.add_trace(go.Scatter(name=str(model), x=m_data[base_col], y=m_data[metric], mode='lines+markers+text', marker=dict(size=8, color=colors[i%len(colors)]), line=dict(width=3, color=colors[i%len(colors)]), text=tv, textposition='top center'))
        else:
            tv = df_grouped[metric].apply(get_tv)
            if is_single_x: 
                fig.add_trace(go.Bar(name=metric, x=df_grouped[base_col], y=df_grouped[metric], marker_color=base_color, text=tv, textposition='inside', insidetextanchor='middle'))
            else: 
                fig.add_trace(go.Scatter(name=metric, x=df_grouped[base_col], y=df_grouped[metric], mode='lines+markers+text', marker=dict(size=8, color=base_color), line=dict(width=3, color=base_color), text=tv, textposition='top center'))
                
        y_range = [0, df_grouped[metric].max() * 1.2 + 2] if not df_grouped.empty and df_grouped[metric].max() > 0 else [0, 10]
        
        fig.update_layout(
            title=f"<b>{title}</b>",
            title_font=dict(color='#1e293b', size=16),
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)', 
            font=dict(color='#334155', size=13), 
            yaxis=dict(gridcolor='#cbd5e1', range=y_range, tickformat=".1f", ticksuffix="%"), 
            xaxis=dict(showgrid=False),
            margin=dict(t=50, b=50, l=10, r=10), 
            legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5), 
            barmode='group' if is_single_x else None, 
            hovermode="x unified",
            hoverlabel=dict(bgcolor="rgba(255,255,255,0.9)", font_size=14, font_color="black", bordercolor="#d3d3d3")
        )
        return fig

    st.markdown("---")
    row1_c1, row1_c2, row1_c3 = st.columns(3)
    with row1_c1: st.plotly_chart(create_single_chart("양품율 트렌드", "양품률", c_yield1), use_container_width=True)
    with row1_c2: st.plotly_chart(create_single_chart("양품율(전,배 포함) 트렌드", "양품율(전/배 포함)", c_yield1), use_container_width=True)
    with row1_c3: st.plotly_chart(create_single_chart("불량율 트렌드", "불량율", c_comp), use_container_width=True)
        
    row2_c1, row2_c2, row2_c3, row2_c4 = st.columns(4)
    with row2_c1: st.plotly_chart(create_single_chart("완전불량율 트렌드", "완전불량률", c_comp), use_container_width=True)
    with row2_c2: st.plotly_chart(create_single_chart("전면불량율 트렌드", "전면불량률", c_front), use_container_width=True)
    with row2_c3: st.plotly_chart(create_single_chart("배면불량율 트렌드", "배면불량률", c_rear), use_container_width=True)
    with row2_c4: st.plotly_chart(create_single_chart("옵셋불량율 트렌드", "옵셋불량율", c_offset), use_container_width=True)

# ==========================================
# 데이터 입력 화면 
# ==========================================
if st.session_state.current_page == "input":

    if st.session_state.get("clear_lot", False):
        st.session_state.lot_input_field = ""
        st.session_state.clear_lot = False
        
    if "temp_lot" in st.session_state:
        st.session_state.lot_input_field = st.session_state.temp_lot
        del st.session_state["temp_lot"]
        
    if "temp_date" in st.session_state:
        st.session_state.in_date_field = st.session_state.temp_date
        del st.session_state["temp_date"]
    
    if "comp_warned" not in st.session_state: st.session_state.comp_warned = False
    if "front_warned" not in st.session_state: st.session_state.front_warned = False
    if "rear_warned" not in st.session_state: st.session_state.rear_warned = False
    if "offset_warned" not in st.session_state: st.session_state.offset_warned = False

    st.markdown("""
        <div style='background: linear-gradient(135deg, #0f172a 0%, #020617 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1e293b;'>
            <h2 style='color: #f8fafc; margin: 0; font-weight: 600;'>VISION DATA KEY-IN SYSTEM</h2>
        </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        # 1. 작업 등록
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.4); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color: #f9fafb; font-weight: 500;'>작업 등록</h4>
            </div>
        """, unsafe_allow_html=True)
        
        default_time = datetime.now().time()
        work_date = st.date_input("**근무일자**", value=datetime.now())
        
        st.markdown("<br><b>교대</b>", unsafe_allow_html=True)
        render_grid_buttons(["주간", "야간"], "shift_type", 2)
        shift_type = st.session_state.shift_type
        
        model_name = st.selectbox("**모델명**", ["D65S(KRIOS)", "MEM", "Centaur", "Sphinx-E", "Banff", "AV-J", "Seattle", "Juliet-O"])
        
        # 2. LOT NO.
        st.markdown("""
            <br>
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.4); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color: #f9fafb; font-weight: 500;'>LOT NO.</h4>
            </div>
        """, unsafe_allow_html=True)
        
        lot_number = st.text_input("**LOT 입력**", placeholder="직접 입력 또는 아래 버튼 스캔", key="lot_input_field")
        
        if st.button("카메라 촬영 및 자동 분석", use_container_width=True):
            open_camera_qr_scanner()
            
        st.markdown("<br>", unsafe_allow_html=True)
        in_date = st.date_input("**입고일**", key="in_date_field")

        # 3. 작업 정보
        st.markdown("""
            <br>
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 15px; border-radius: 8px; margin-bottom: 10px; box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.4); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color: #f9fafb; font-weight: 500;'>작업 정보</h4>
            </div>
        """, unsafe_allow_html=True)
        
        start_date = st.date_input("**시작일**", value=datetime.now())
        start_time = st.time_input("**시작시간**", value=default_time, key="start_time_key")
        
        end_date = st.date_input("**종료일**", value=datetime.now())
        end_time = st.time_input("**종료시간**", value=default_time, key="end_time_key")
        
        st.markdown("<br><b>호기</b>", unsafe_allow_html=True)
        render_grid_buttons(["1호기", "2호기", "3호기", " "], "unit", 2)
        unit = st.session_state.unit
        
        st.markdown("<br><b>검사 구분</b>", unsafe_allow_html=True)
        render_grid_buttons(["1차 검사", "2차 검사"], "category", 2)
        category = st.session_state.category
        
        st.markdown("<br><b>도금 구분</b>", unsafe_allow_html=True)
        render_grid_buttons(["A", "B"], "plating_type", 2)
        plating_type = st.session_state.plating_type
        
        st.markdown("<br>", unsafe_allow_html=True)
        idle_time = st.number_input("**휴동시간 (분)**", min_value=0, value=0)
        
        start_dt = datetime.combine(start_date, start_time)
        end_dt = datetime.combine(end_date, end_time)
        raw_duration = int((end_dt - start_dt).total_seconds() / 60)
        duration_minutes = max(0, raw_duration - idle_time)
            
        st.text_input("**소요시간 (휴동시간 차감됨)**", value=f"{duration_minutes:,} 분", disabled=True)

        # 백업 영역
        st.markdown("<br><br><hr>", unsafe_allow_html=True)
        st.write("백업 및 보관용")
        export_df = load_data().copy()
        if not export_df.empty:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                export_df.to_excel(writer, index=False)
                wb = writer.book
                ws = wb.active
                red_font = Font(color="FF0000")
                headers = {cell.value: i for i, cell in enumerate(ws[1])}
                for row in ws.iter_rows(min_row=2):
                    for c_name in ["UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타", "소요시간", "휴동시간"]:
                        if c_name in headers: row[headers[c_name]].number_format = '#,##0'
                    try: yield_val = float(str(row[headers['양품률']].value).replace('%', '').strip())
                    except: yield_val = 100.0
                    if yield_val < 85.0 and '양품률' in headers: row[headers['양품률']].font = red_font
            st.download_button(label="DB 데이터를 엑셀로 다운로드", data=output.getvalue(), file_name=f"VISION_EXPORT_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        st.markdown("<br><hr>", unsafe_allow_html=True)
        with st.expander("외부 엑셀 데이터 대량 업로드", expanded=False):
            uploaded_file = st.file_uploader("엑셀 선택", type=['xlsx'])
            if uploaded_file and st.button("스마트 분석 및 구글 DB 저장", type="primary", use_container_width=True):
                with st.spinner("저장 중..."):
                    xls = pd.ExcelFile(uploaded_file)
                    target_sheet = next((s for s in xls.sheet_names if 'Q' in s or '년' in s), xls.sheet_names[0])
                    temp_df = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=100)
                    target_header_idx = next((idx for idx, row in temp_df.iterrows() if sum([1 for kw in ['날짜', '교대', '모델'] if kw in str(row).replace(' ', '')]) >= 2), None)
                    if target_header_idx is not None:
                        import_df = pd.read_excel(xls, sheet_name=target_sheet, header=target_header_idx)
                        def smart_map(c):
                            c=str(c).replace('\n','').replace(' ','')
                            if '모델' in c: return '모델명(MI)'
                            if '날짜' in c or '일자' in c: return '날짜'
                            if '교대' in c: return '교대'
                            if '시작' in c: return '시작시간'
                            if '종료' in c: return '종료시간'
                            if 'LOT' in c.upper(): return 'LOT NO.'
                            if '육안' in c or 'OQC' in c.upper(): return 'OQC'
                            if '도장라인' in c or 'LINE' in c.upper(): return '도장라인'
                            if '검사' in c: return '검사 수량'
                            if '양품' in c and ('전/배' in c or '전,배' in c):
                                if '율' in c or '률' in c: return '양품율(전/배 포함)'
                                else: return '양품 수량(전/배 포함)'
                            if '양품' in c:
                                if '율' in c or '률' in c: return '양품률'
                                else: return '양품수량'
                            if '불량수량' in c or ('불량' in c and '수량' in c): return '불량수량'
                            if '완전' in c: return '완전불량'
                            if '전면' in c and '율' not in c: return '전면불량'
                            if '배면' in c and '율' not in c: return '배면불량'
                            if '옵셋' in c: return '옵셋불량'
                            return c
                        import_df.columns = [smart_map(c) for c in import_df.columns]
                        import_df = import_df.dropna(subset=['날짜', '모델명(MI)'])
                        import_df['시작시간'] = import_df['종료시간'] = ""
                        import_df['휴동시간'] = import_df['소요시간'] = import_df['UPH'] = import_df['UPD'] = 0
                        for col in EXCEL_COLUMNS:
                            if col not in import_df.columns: import_df[col] = 0 if col in ["검사 수량"] else ""
                        if save_data_append(import_df[EXCEL_COLUMNS]): st.success("저장 성공!")
                    else: st.error("제목줄을 찾을 수 없습니다.")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("종합 분석 데이터 확인", use_container_width=True):
            show_password_dialog()
            
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 최신 데이터 새로고침", use_container_width=True):
            load_data.clear()
            st.rerun()

    main_col1, main_col2 = st.columns([1.1, 0.9])
    save_success_trigger = False

    with main_col1:
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>📥 VISION Data</h4>
            </div>
        """, unsafe_allow_html=True)
        
        v_row1_col1, v_row1_col2, v_row1_col3 = st.columns(3)
        v_row2_col1, v_row2_col2, v_row2_col3 = st.columns(3)
        v_row3_col1, v_row3_col2, v_row3_col3 = st.columns(3)
        
        with v_row1_col2: good_qty = st.number_input("**양품수량**", min_value=0, value=0)
        with v_row2_col1: comp_def = st.number_input("**완전불량**", min_value=0, value=0)
        with v_row2_col2: front_def = st.number_input("**전면불량**", min_value=0, value=0)
        with v_row2_col3: rear_def = st.number_input("**배면불량**", min_value=0, value=0)
        with v_row3_col1: offset_def = st.number_input("**옵셋불량**", min_value=0, value=0)
        with v_row3_col2: shortage_qty = st.number_input("**수량부족**", min_value=0, value=0)
        with v_row3_col3: etc_def = st.number_input("**기타**", min_value=0, value=0)
        
        bad_qty = comp_def + front_def + rear_def + offset_def + etc_def
        total_qty = max(0, good_qty + bad_qty - shortage_qty)
            
        with v_row1_col1: st.text_input("**검사 수량 (자동)**", value=f"{total_qty:,}", disabled=True)
        with v_row1_col3: st.text_input("**불량수량 (자동)**", value=f"{bad_qty:,}", disabled=True)

        uph_val = int((total_qty / duration_minutes) * 60) if duration_minutes > 0 else 0
        upd_val = uph_val * 22

        good_include_front_rear = good_qty + front_def + rear_def
        if total_qty > 0:
            rate_good = round((good_qty / total_qty) * 100, 1)
            rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
            rate_front = round((front_def / total_qty) * 100, 1)
            rate_rear = round((rear_def / total_qty) * 100, 1)
            rate_bad = round((bad_qty / total_qty) * 100, 1)
            comp_rate_num = round(comp_def / total_qty * 100, 1)
            front_rate_num = round(front_def / total_qty * 100, 1)
            rear_rate_num = round(rear_def / total_qty * 100, 1)
            offset_rate_num = round(offset_def / total_qty * 100, 1)
        else:
            rate_good = rate_good_inc = rate_front = rate_rear = rate_bad = 0.0
            comp_rate_num = front_rate_num = rear_rate_num = offset_rate_num = 0.0

        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>Coating Data</h4>
            </div>
        """, unsafe_allow_html=True)
        
        coat_col1, coat_col2, coat_col3 = st.columns([1, 1, 2])
        with coat_col1: painting_date = st.date_input("**도장일**")
        with coat_col2: painting_order = st.number_input("**도장순서**", min_value=1, value=1)
        with coat_col3:
            st.markdown("<b>도장라인</b>", unsafe_allow_html=True)
            render_grid_buttons(["A Line", "B Line", "C Line"], "painting_line", 3)
            painting_line = st.session_state.painting_line
            
        st.markdown("""
            <br>
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>Assemble Data</h4>
            </div>
        """, unsafe_allow_html=True)
        
        parts_col1, parts_col2, parts_col3 = st.columns(3)
        num_options = ["선택안함"] + [str(i) for i in range(1, 11)]
        with parts_col1: clip_val = st.selectbox("**CLIP**", num_options, index=1)
        with parts_col2: base_val = st.selectbox("**BASE**", num_options, index=1)
        with parts_col3: cover_val = st.selectbox("**COVER**", num_options, index=1)

        st.markdown("<br><b>조립기</b>", unsafe_allow_html=True)
        render_grid_buttons(["1호기", "2호기", "3호기", "4호기"], "assembler_val", 2)
        assembler_val = st.session_state.assembler_val

        st.markdown("""
            <br>
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>ETC.</h4>
            </div>
        """, unsafe_allow_html=True)

        qc_col1, qc_col2 = st.columns(2)
        with qc_col1: oqc_status = st.selectbox("**OQC**", ["선택안함", "육안", "OQC"])
        with qc_col2: worker_name = st.selectbox("**작업자**", ["선택안함"] + worker_list)
        
        remarks = st.text_area("**비고**", height=68, placeholder="특이사항을 입력하세요.")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("데이터 저장 (구글 스프레드시트)", use_container_width=True, type="primary"):
            if total_qty == 0: st.warning("입력된 데이터가 없습니다.")
            elif not lot_number: st.warning("LOT 번호를 입력하거나 스캔해주세요.")
            else:
                with st.spinner("저장 및 중복 검증 중..."):
                    db_df = load_data()
                    is_dup = False
                    if not db_df.empty and "LOT NO." in db_df.columns and "교대" in db_df.columns:
                        if len(db_df[(db_df["LOT NO."] == lot_number) & (db_df["교대"] == shift_type)]) > 0:
                            is_dup = True

                    if is_dup: st.error("이미 처리되었습니다. (동일 교대 내 중복 LOT)")
                    else:
                        fmt_date = f"{work_date.month}/{work_date.day}"
                        fmt_paint_date = f"{painting_date.month}/{painting_date.day}" 
                        fmt_in_date = in_date.strftime("%Y-%m-%d") 
                        
                        fmt_paint_line = painting_line.replace(" Line", "") if painting_line != "선택안함" else ""
                        fmt_assembler = assembler_val.replace("호기", "") if assembler_val != "선택안함" else ""
                        fmt_oqc = "" if oqc_status == "선택안함" else oqc_status
                        fmt_worker = "" if worker_name == "선택안함" else worker_name
                        fmt_clip = "" if clip_val == "선택안함" else clip_val
                        fmt_base = "" if base_val == "선택안함" else base_val
                        fmt_cover = "" if cover_val == "선택안함" else cover_val
                        fmt_lot = lot_number 
                        
                        new_data = pd.DataFrame([{
                            "날짜": fmt_date, "교대": shift_type,
                            "시작시간": start_dt.strftime("%H:%M"), "종료시간": end_dt.strftime("%H:%M"),
                            "휴동시간": f"{idle_time:,}", "소요시간": f"{duration_minutes:,}", "구분": category, "호기": unit, 
                            "모델명(MI)": model_name, "도금구분": plating_type, "UPH": f"{uph_val:,}", "UPD": f"{upd_val:,}",
                            "검사 수량": f"{total_qty:,}", "양품수량": f"{good_qty:,}", "양품 수량(전/배 포함)": f"{good_include_front_rear:,}", "불량수량": f"{bad_qty:,}",
                            "양품률": f"{rate_good:.1f}%", "양품율(전/배 포함)": f"{rate_good_inc:.1f}%",
                            "완전불량률": f"{comp_rate_num:.1f}%", "전면불량률": f"{front_rate_num:.1f}%", "배면불량률": f"{rear_rate_num:.1f}%",
                            "완전불량": f"{comp_def:,}", "전면불량": f"{front_def:,}", "배면불량": f"{rear_def:,}", "옵셋불량": f"{offset_def:,}", "수량부족": f"{shortage_qty:,}", "기타": f"{etc_def:,}",
                            "OQC": fmt_oqc, "비고": remarks, "도장라인": fmt_paint_line, "도장일": fmt_paint_date, 
                            "도장순서": painting_order, "입고일": fmt_in_date, "LOT NO.": fmt_lot, 
                            "CLIP": fmt_clip, "BASE": fmt_base, "COVER": fmt_cover, "조립기": fmt_assembler, 
                            "월": f"{work_date.month}월", "작업자": fmt_worker
                        }])
                        if save_data_append(new_data):
                            st.success("저장 완료!")
                            st.session_state.clear_lot = True
                            save_success_trigger = True  

    if category == "1차 검사" and total_qty > 0:
        if comp_rate_num > 5.0:
            if not st.session_state.comp_warned:
                show_sbl_warning("완전불량", comp_rate_num)
                st.session_state.comp_warned = True
        else: st.session_state.comp_warned = False
        if front_rate_num > 5.0:
            if not st.session_state.front_warned:
                show_sbl_warning("전면불량", front_rate_num)
                st.session_state.front_warned = True
        else: st.session_state.front_warned = False
        if rear_rate_num > 5.0:
            if not st.session_state.rear_warned:
                show_sbl_warning("배면불량", rear_rate_num)
                st.session_state.rear_warned = True
        else: st.session_state.rear_warned = False
        if offset_rate_num > 5.0:
            if not st.session_state.offset_warned:
                show_sbl_warning("옵셋불량", offset_rate_num)
                st.session_state.offset_warned = True
        else: st.session_state.offset_warned = False
    else:
        st.session_state.comp_warned = st.session_state.front_warned = st.session_state.rear_warned = st.session_state.offset_warned = False

    with main_col2:
        st.markdown("""
            <div style='background: linear-gradient(135deg, #111827 0%, #030712 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1f2937;'>
                <h4 style='margin: 0; color:#f9fafb; font-weight: 500;'>Yield Report</h4>
            </div>
        """, unsafe_allow_html=True)
        
        m_col1, m_col2 = st.columns(2)
        with m_col1: st.metric(label="검사수량 총합", value=f"{total_qty:,} EA")
        with m_col2: st.metric(label="현재 양품율", value=f"{rate_good:.1f}%")

        with st.expander("🎨 Chart Option (색상 및 폰트 설정)", expanded=False):
            st.markdown("**폰트 크기 설정**")
            chart_font_size = st.slider("그래프 폰트 텍스트 크기", min_value=10, max_value=30, value=16)

            st.markdown("**그래프 색상 설정**")
            color_col1, color_col2, color_col3, color_col4, color_col5 = st.columns(5)
            with color_col1: c_yield = st.color_picker("양품", "#002b5e") 
            with color_col2: c_comp  = st.color_picker("완전불량", "#b22222")
            with color_col3: c_front = st.color_picker("전면불량", "#ed7d31")
            with color_col4: c_rear  = st.color_picker("배면불량", "#00b050")
            with color_col5: c_offset= st.color_picker("옵셋불량", "#7030a0")

        fig_donut = go.Figure(go.Pie(
            labels=['양품율', '불량율'], 
            values=[rate_good, rate_bad], 
            hole=.65, 
            sort=False,
            direction='clockwise',
            marker=dict(
                colors=[c_yield, '#e2e8f0'], 
                line=dict(color='#ffffff', width=2)
            ), 
            hoverinfo="label+percent", 
            textinfo="none"
        ))
        
        fig_donut.update_layout(
            showlegend=False,
            height=300,
            margin=dict(t=20, b=10, l=10, r=10),
            paper_bgcolor='rgba(0,0,0,0)', 
            plot_bgcolor='rgba(0,0,0,0)',
            annotations=[
                dict(text="현재 양품율", x=0.5, y=0.62, font_size=15, font_color="#475569", showarrow=False),
                dict(text=f"{rate_good:.1f}%", x=0.5, y=0.45, font_size=chart_font_size * 2.8, font_color=c_yield, font_family="Arial, sans-serif", font_weight="bold", showarrow=False)
            ]
        )
        st.plotly_chart(fig_donut, use_container_width=True)
        
        df_defects = pd.DataFrame({"불량 항목": ['완전불량', '전면불량', '배면불량', '옵셋불량'], "비율 (%)": [comp_rate_num, front_rate_num, rear_rate_num, offset_rate_num]})
        y_max = max(df_defects["비율 (%)"]) * 1.4 if not df_defects.empty and max(df_defects["비율 (%)"]) > 0 else 5

        fig_bar = go.Figure()
        colors = [c_comp, c_front, c_rear, c_offset]
        
        fig_bar.add_trace(go.Bar(
            x=df_defects["불량 항목"],
            y=[y_max] * 4,
            marker_color='#f1f5f9',
            hoverinfo='none',
            width=0.45
        ))

        fig_bar.add_trace(go.Bar(
            x=df_defects["불량 항목"],
            y=df_defects["비율 (%)"],
            marker_color=colors,
            width=0.45,
            texttemplate='',
        ))

        dynamic_marker_size = chart_font_size * 2.5 

        fig_bar.add_trace(go.Scatter(
            x=df_defects["불량 항목"],
            y=df_defects["비율 (%)"],
            mode='markers+text',
            marker=dict(
                size=dynamic_marker_size,
                color=colors,
                line=dict(color='white', width=3) 
            ),
            text=df_defects["비율 (%)"].apply(lambda x: f"{x:.1f}"),
            textfont=dict(color='white', size=chart_font_size, weight='bold'),
            textposition='middle center',
            hoverinfo='none'
        ))

        fig_bar.update_layout(
            barmode='overlay', 
            showlegend=False,
            height=300,
            margin=dict(t=10, b=20, l=10, r=10),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(showgrid=False, tickfont=dict(color='#334155', size=14, weight='bold')),
            yaxis=dict(showgrid=False, showticklabels=False, range=[0, y_max]) 
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("<br>", unsafe_allow_html=True)
        logo_c1, logo_c2, logo_c3 = st.columns([1, 1, 1])
        with logo_c2:
            try:
                st.image("at.png", use_container_width=True)
            except Exception:
                pass 

    st.markdown("---") 
    with st.expander("최근 저장 데이터 List (구글 시트 연동중 - 수정 가능)", expanded=True):
        
        sheet_info = get_sheet()
        if sheet_info:
            st.caption(f"🔗 연결된 구글 시트 탭: **[{sheet_info.title}]** (데이터가 안 보인다면 구글 시트의 **{sheet_info.title}** 탭에 알맞은 데이터가 있는지 확인해 주세요!)")
            
        df_history = load_data().copy()
        if not df_history.empty:
            recent_10 = df_history.iloc[::-1].head(10).copy()
            valid_cols = [col for col in EXCEL_COLUMNS if col in recent_10.columns]
            
            num_cols = ["UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타", "소요시간", "휴동시간"]
            display_df = recent_10[valid_cols].copy()
            for col in num_cols:
                if col in display_df.columns:
                    display_df[col] = display_df[col].apply(lambda x: f"{float(x):,.0f}" if pd.notnull(x) and str(x).replace('.','',1).isdigit() else "0")
                    
            edited_df = st.data_editor(display_df, use_container_width=True, hide_index=True)
            
            if not display_df.equals(edited_df):
                st.info("데이터가 변경되었습니다. 값을 수정한 후 아래 덮어쓰기 버튼을 눌러주세요.")
                if st.button("변경된 데이터 구글 시트에 덮어쓰기", type="primary"):
                    with st.spinner("구글 스프레드시트에 데이터를 덮어쓰는 중입니다..."):
                        df_full = load_data().copy()
                        changes_applied = False
                        
                        for idx in edited_df.index:
                            orig_row = display_df.loc[idx]
                            edit_row = edited_df.loc[idx]
                            
                            if not orig_row.equals(edit_row):
                                def get_val(col_name):
                                    try: return int(str(edit_row.get(col_name, 0)).replace(',', ''))
                                    except: return 0
                                    
                                good_qty = get_val("양품수량")
                                comp_def = get_val("완전불량")
                                front_def = get_val("전면불량")
                                rear_def = get_val("배면불량")
                                offset_def = get_val("옵셋불량")
                                shortage_qty = get_val("수량부족")
                                etc_def = get_val("기타")
                                
                                bad_qty = comp_def + front_def + rear_def + offset_def + etc_def
                                total_qty = good_qty + bad_qty - shortage_qty
                                if total_qty < 0: total_qty = 0
                                good_include_front_rear = good_qty + front_def + rear_def

                                dur_min = get_val("소요시간")
                                uph_update = int((total_qty / dur_min) * 60) if dur_min > 0 else 0
                                upd_update = uph_update * 22
                                
                                if total_qty > 0:
                                    rate_good = round((good_qty / total_qty) * 100, 1)
                                    rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
                                    comp_rate_num = round(comp_def / total_qty * 100, 1)
                                    front_rate_num = round(front_def / total_qty * 100, 1)
                                    rear_rate_num = round(rear_def / total_qty * 100, 1)
                                else:
                                    rate_good = rate_good_inc = comp_rate_num = front_rate_num = rear_rate_num = 0.0
                                    
                                for col in valid_cols:
                                    if col == "검사 수량": val = f"{total_qty:,}"
                                    elif col == "불량수량": val = f"{bad_qty:,}"
                                    elif col == "양품 수량(전/배 포함)": val = f"{good_include_front_rear:,}"
                                    elif col == "양품률": val = f"{rate_good:.1f}%"
                                    elif col == "양품율(전/배 포함)": val = f"{rate_good_inc:.1f}%"
                                    elif col == "완전불량률": val = f"{comp_rate_num:.1f}%"
                                    elif col == "전면불량률": val = f"{front_rate_num:.1f}%"
                                    elif col == "배면불량률": val = f"{rear_rate_num:.1f}%"
                                    elif col == "UPH": val = f"{uph_update:,}"
                                    elif col == "UPD": val = f"{upd_update:,}"
                                    elif col == "LOT NO.":
                                        val = str(edit_row.get(col, ""))
                                    elif col in num_cols: val = f"{get_val(col):,}"
                                    else: val = edit_row[col]
                                    df_full.at[idx, col] = val
                                changes_applied = True
                                
                        if changes_applied:
                            if save_data_overwrite(df_full[valid_cols]):
                                st.success("변경된 데이터가 구글 스프레드시트에 완벽하게 덮어씌워졌습니다!")
                                import time
                                time.sleep(1)
                                st.rerun()
        else:
            st.caption("현재 연결된 탭에 저장된 데이터가 없거나, 인식할 수 없습니다.")

    if save_success_trigger:
        import time
        time.sleep(1)
        st.rerun()

# ==========================================
# 종합 데이터 분석 화면 렌더링
# ==========================================
elif st.session_state.current_page == "analysis":
    render_analysis_page()