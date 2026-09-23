import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import json
import os
import re
import base64
import uuid
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="VISION DATA KEY-IN SYSTEM", layout="wide", initial_sidebar_state="collapsed")

# ==========================================
# 💡 헬퍼 함수 모음
# ==========================================
def hex_to_rgba(hex_color, alpha):
    hex_color = hex_color.lstrip('#')
    hlen = len(hex_color)
    rgb = tuple(int(hex_color[i:i+hlen//3], 16) for i in range(0, hlen, hlen//3))
    return f"rgba({rgb[0]},{rgb[1]},{rgb[2]},{alpha})"

def get_image_base64(base_name):
    try:
        extensions = ['.png', '.jpg', '.jpeg']
        search_dirs = [os.getcwd(), os.path.dirname(os.path.abspath(__file__))]
        for directory in search_dirs:
            for ext in extensions:
                filepath = os.path.join(directory, base_name + ext)
                if os.path.exists(filepath):
                    with open(filepath, "rb") as img_file:
                        encoded = base64.b64encode(img_file.read()).decode('utf-8')
                        mime_type = "image/jpeg" if ext in ['.jpg', '.jpeg'] else "image/png"
                        return f"data:{mime_type};base64,{encoded}"
    except Exception:
        pass
    return None

SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
SPREADSHEET_ID = "1DeMJJkuq7bYa4XNK_NbkqZ-vOJKqGhmYXIvHm3yJl8E"
TAB_NAME = "2026년 3Q"
EXCEL_COLUMNS = [
    "고유 ID", "상태", "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", 
    "모델명(MI)", "도금구분", "UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", 
    "불량수량", "양품율", "양품율(전/배 포함)", "완전불량율", "전면불량율", "배면불량율", 
    "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타", "OQC", "비고", 
    "도장라인", "도장일", "도장순서", "입고일", "LOT NO.", "CLIP", "BASE", "COVER", 
    "조립기", "월", "작업자"
]

@st.cache_resource(ttl=600)
def get_spreadsheet_doc():
    try:
        creds_data = st.secrets["google_credentials"]
        clean_data = creds_data.strip().strip("'").strip('"') if isinstance(creds_data, str) else dict(creds_data)
        creds_dict = json.loads(clean_data, strict=False) if isinstance(creds_data, str) else clean_data
        if "private_key" in creds_dict: creds_dict["private_key"] = creds_dict["private_key"].replace('\\n', '\n')
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
        doc = gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
        return doc
    except Exception as e:
        st.error(f"🚨 구글 API 인증/연결 오류: {e}")
        return None

def load_shared_config():
    doc = get_spreadsheet_doc()
    if doc:
        try:
            ws = doc.worksheet("VIEWER_CONFIG")
            val = ws.acell('A1').value
            if val: return json.loads(val)
        except Exception:
            return None
    return None

def save_shared_config(config_dict):
    doc = get_spreadsheet_doc()
    if doc:
        try:
            ws = doc.worksheet("VIEWER_CONFIG")
            ws.update_acell('A1', json.dumps(config_dict, ensure_ascii=False))
            st.success("✅ 시스템 설정이 클라우드에 성공적으로 반영되었습니다.")
            return True
        except Exception as e:
            st.error(f"🚨 설정 저장 실패: {e}")
            return False
    return False

@st.cache_data(ttl=15)
def load_universal_data():
    doc = get_spreadsheet_doc()
    if doc is None: return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row', 'DateTime', 'DateOnly'])
    try:
        ws = doc.worksheet(TAB_NAME)
        raw_data = ws.get_all_values()
    except Exception as e:
        st.error(f"🚨 '{TAB_NAME}' 시트 접근 에러: {e}")
        return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row', 'DateTime', 'DateOnly'])
    
    if len(raw_data) < 24: 
        return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row', 'DateTime', 'DateOnly'])
    
    header_idx = 22
    data_start_idx = 23
    raw_headers = [str(h).strip().replace('\n', '') for h in raw_data[header_idx]]
    
    unique_headers = []
    seen = {}
    for h in raw_headers:
        h_clean = h if h != "" else "UNNAMED"
        if h_clean in seen:
            seen[h_clean] += 1
            unique_headers.append(f"{h_clean}_{seen[h_clean]}")
        else:
            seen[h_clean] = 0
            unique_headers.append(h_clean)
            
    df = pd.DataFrame(raw_data[data_start_idx:], columns=unique_headers, dtype=str)
    
    rename_dict = {}
    mapped_std_cols = set()
    for c_idx, h in enumerate(unique_headers):
        cc = h.replace(" ", "").replace("률", "율").upper()
        matched_col = None
        if c_idx == 2: matched_col = '날짜'
        elif "고유" in cc and "ID" in cc: matched_col = '고유 ID'
        elif cc == "상태": matched_col = '상태'
        elif cc == "시작시간": matched_col = '시작시간'
        elif cc == "소요시간": matched_col = '소요시간'
        elif cc == "구분": matched_col = '구분'
        elif cc in ["모델명(MI)", "모델명", "품명", "MI"]: matched_col = '모델명(MI)'
        elif cc in ["검사수량", "총수량"]: matched_col = '검사 수량'
        elif cc == "양품수량": matched_col = '양품수량'
        elif "양품수량" in cc and "포함" in cc: matched_col = '양품 수량(전/배 포함)'
        elif cc in ["양품율", "수율", "합격율"]: matched_col = '양품율'
        elif ("양품율" in cc or "수율" in cc) and "포함" in cc: matched_col = '양품율(전/배 포함)'
        elif cc in ["완전불량율", "완전불량률"]: matched_col = '완전불량율'
        elif cc in ["전면불량율", "전면불량률"]: matched_col = '전면불량율'
        elif cc in ["배면불량율", "배면불량률"]: matched_col = '배면불량율'
        elif cc == "완전불량": matched_col = '완전불량'
        elif cc == "전면불량": matched_col = '전면불량'
        elif cc == "배면불량": matched_col = '배면불량'
        elif cc == "옵셋불량": matched_col = '옵셋불량'
        elif cc == "옵셋" in cc and ("율" in cc or "률" in cc): matched_col = '옵셋불량율' 
        elif cc == "도장일": matched_col = '도장일'
        elif cc == "도장순서": matched_col = '도장순서'
        elif cc in ["LOTNO.", "LOTNO", "LOT", "로트"]: matched_col = 'LOT NO.'
        elif "날짜" in cc and c_idx != 2: matched_col = '날짜'

        if matched_col and matched_col not in mapped_std_cols:
            rename_dict[h] = matched_col
            mapped_std_cols.add(matched_col)

    df = df.rename(columns=rename_dict)
    ext_cols = EXCEL_COLUMNS + ['옵셋불량율']
    for col in ext_cols:
        if col not in df.columns: df[col] = ""
        
    def parse_dt(r):
        d_val = str(r.get('날짜', '')).strip()
        t_val = str(r.get('시작시간', '00:00')).strip()
        if not d_val or d_val.lower() in ['nan', 'none']: return datetime(2026, 1, 1) 
        t_clean = re.sub(r'[^\d]', '', str(t_val))
        if len(t_clean) >= 4: t_str = f"{t_clean[:2]}:{t_clean[2:4]}:00"
        elif len(t_clean) == 3: t_str = f"0{t_clean[:1]}:{t_clean[1:3]}:00"
        elif len(t_clean) in [1, 2]: t_str = f"{t_clean.zfill(2)}:00:00"
        else: t_str = "00:00:00"
        try:
            if d_val.isdigit() and 40000 <= int(d_val) <= 50000:
                return pd.to_datetime(f"{(datetime(1899, 12, 30) + timedelta(days=int(d_val))).strftime('%Y-%m-%d')} {t_str}", errors='coerce') or datetime(2026, 1, 1)
            parts = re.split(r'[./-]', d_val)
            if len(parts) == 3:
                p1, p2, p3 = int(parts[0]), int(parts[1]), int(parts[2])
                if p1 > 1000: return pd.to_datetime(f"{p1}-{p2:02d}-{p3:02d} {t_str}", errors='coerce') or datetime(2026, 1, 1)
                elif p3 > 1000: return pd.to_datetime(f"{p3}-{p1:02d}-{p2:02d} {t_str}", errors='coerce') or datetime(2026, 1, 1)
                else: return pd.to_datetime(f"20{p3:02d}-{p1:02d}-{p2:02d} {t_str}", errors='coerce') or datetime(2026, 1, 1)
            if len(parts) == 2:
                return pd.to_datetime(f"2026-{int(parts[0]):02d}-{int(parts[1]):02d} {t_str}", errors='coerce') or datetime(2026, 1, 1)
        except: pass
        return datetime(2026, 1, 1) 
        
    if df.empty:
        df['DateTime'] = pd.to_datetime([])
        df['DateOnly'] = []
    else:
        parsed_dates = df.apply(parse_dt, axis=1)
        missing_dates_idx = parsed_dates.isna()
        if missing_dates_idx.any():
            parsed_dates.loc[missing_dates_idx] = [datetime(2026, 1, 1) + timedelta(minutes=i) for i in range(missing_dates_idx.sum())]
        df['DateTime'] = pd.to_datetime(parsed_dates)
        df['DateOnly'] = df['DateTime'].dt.date 

    if '구분' in df.columns:
        df_filtered = df[df['구분'].fillna('').astype(str).str.contains('1차', na=False)]
        if not df_filtered.empty: df = df_filtered
        
    final_cols = ext_cols + ['DateTime', 'DateOnly']
    for c in final_cols:
        if c not in df.columns:
            df[c] = None
            
    return df[final_cols]

# ==========================================
# 💡 시스템 상태 변수 초기화
# ==========================================
if "main_authenticated" not in st.session_state: st.session_state.main_authenticated = False
if "sys_menu" not in st.session_state: st.session_state.sys_menu = "keyin"
if "rotate_idx" not in st.session_state: st.session_state.rotate_idx = 0

# ==========================================
# 💡 전역 CSS 및 뱃지 철통 방어 스크립트 (모든 화면 공통)
# ==========================================
global_theme_css = """
<style>
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { display: none !important; visibility: hidden !important; }
[data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
footer { display: none !important; } 

[data-testid="collapsedControl"] { display: none !important; pointer-events: none !important; }
[data-testid="stSidebar"] { display: none !important; }
body { overscroll-behavior-y: none !important; background-color: #f8fafc !important; } 
::-webkit-scrollbar { display: none; }
.block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 98% !important; }

h1, h2, h3, h4, h5, h6, p, label { font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif !important; color: #1e293b !important; }
[data-testid="stAppViewContainer"] { background-color: #f8fafc !important; color: #1e293b !important; }
div[data-baseweb="input"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }
div[data-baseweb="input"] input { color: #1e293b !important; font-weight: bold !important; }

div[data-testid="stVerticalBlockBorderWrapper"] { background-color: #ffffff !important; border-radius: 12px !important; border: 1px solid #e2e8f0 !important; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03) !important; padding: 1.5rem !important; margin-bottom: 0.8rem !important; }
.command-header { color: #1e293b !important; font-weight: 900 !important; letter-spacing: 1px; }

/* 공통 버튼 스타일 */
div[data-testid="stButton"] button { height: 2.6rem !important; min-height: 2.6rem !important; font-size: 1.1rem !important; font-weight: bold !important; border-radius: 8px !important; background-color: #E7E6E6 !important; border: 1px solid #cbd5e1 !important; transition: all 0.2s ease; }
div[data-testid="stButton"] button p { color: #000000 !important; }
div[data-testid="stButton"] button:hover { background-color: #1e293b !important; border-color: #1e293b !important; }
div[data-testid="stButton"] button:hover p { color: #ffffff !important; }

div[data-testid="stButton"] button[kind="primary"] { background-color: #1e293b !important; border: 1px solid #0f172a !important; }
div[data-testid="stButton"] button[kind="primary"]:hover { background-color: #0f172a !important; }
div[data-testid="stButton"] button[kind="primary"] p { color: #ffffff !important; }
</style>
"""
st.markdown(global_theme_css, unsafe_allow_html=True)

badge_killer_script = """
<script>
const setupBadgeBlocker = () => {
    let docs = [document];
    try { if (window.parent && window.parent.document) docs.push(window.parent.document); } catch(e){}
    try { if (window.top && window.top.document && window.top !== window.parent) docs.push(window.top.document); } catch(e){}

    docs.forEach(doc => {
        try {
            const selectors = '[data-testid="manage-app-button"], [data-testid="stAppDeployButton"], .stDeployButton, div[class^="viewerBadge"], div[class*="viewerBadge"], #creatorBadge, a[href*="streamlit.io/cloud"]';
            doc.querySelectorAll(selectors).forEach(el => {
                el.style.setProperty('display', 'none', 'important');
                el.style.setProperty('pointer-events', 'none', 'important');
            });
            
            doc.querySelectorAll('div, a, button, span').forEach(el => {
                if (el.textContent && (el.textContent.includes('< Manage app') || el.textContent.includes('View profile'))) {
                    el.style.setProperty('display', 'none', 'important');
                    if (el.parentElement) el.parentElement.style.setProperty('display', 'none', 'important');
                }
            });

            if (!doc.getElementById('ultimate-blocker-shield')) {
                const blocker = doc.createElement('div');
                blocker.id = 'ultimate-blocker-shield';
                blocker.style.cssText = 'position:fixed !important; bottom:0 !important; right:0 !important; width:300px !important; height:150px !important; background:transparent !important; z-index:2147483647 !important; cursor:default !important; pointer-events:auto !important;';
                
                const killEvent = (e) => { e.stopPropagation(); e.preventDefault(); return false; };
                ['click', 'mousedown', 'mouseup', 'pointerdown', 'touchstart'].forEach(ev => blocker.addEventListener(ev, killEvent, true));
                doc.body.appendChild(blocker);
            }
        } catch(e) {}
    });
};
setupBadgeBlocker();
setInterval(setupBadgeBlocker, 100); 
</script>
"""
components.html(badge_killer_script, height=0, width=0)

# ==========================================
# 💡 로그인 화면
# ==========================================
if not st.session_state.main_authenticated:
    st.markdown("<br><br><br><br><br>", unsafe_allow_html=True)
    col_sp1, col_auth, col_sp3 = st.columns([1, 1, 1])
    with col_auth:
        with st.container(border=True):
            logo_l_data = get_image_base64("logo")
            if logo_l_data:
                st.markdown(f"<div style='text-align: center;'><img src='{logo_l_data}' style='max-width: 100%; max-height: 80px; object-fit: contain; margin-bottom: 15px;'></div>", unsafe_allow_html=True)
                
            st.markdown("<h3 style='text-align:center; color:#1e293b; font-weight:900;'>🔐 COMMAND CENTER 접속 인증</h3>", unsafe_allow_html=True)
            st.markdown("<div style='text-align:center; color:#64748b; margin-bottom:20px; font-weight:bold;'>데이터 입력 및 시스템 접근을 위해 비밀번호를 입력하세요.</div>", unsafe_allow_html=True)
            pwd = st.text_input("비밀번호", type="password", label_visibility="collapsed", placeholder="비밀번호 입력", key="main_pwd")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✅ 접속", type="primary", use_container_width=True, key="main_confirm"):
                if pwd == "7777":  # 메인 시스템 비밀번호 설정
                    st.session_state.main_authenticated = True
                    st.rerun()
                else:
                    st.error("비밀번호가 일치하지 않습니다.")
    st.stop()

# ==========================================
# 💡 최상단 네비게이션 메뉴 (접속 성공 후)
# ==========================================
st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 900; margin-bottom: 20px;'>VISION DATA COMMAND CENTER</h2>", unsafe_allow_html=True)

m_col1, m_col2, m_col3, m_col4 = st.columns([0.28, 0.28, 0.28, 0.16])
with m_col1:
    if st.button("📝 KEY-IN WIZARD", use_container_width=True, type="primary" if st.session_state.sys_menu == "keyin" else "secondary"):
        st.session_state.sys_menu = "keyin"
        st.rerun()
with m_col2:
    if st.button("👁️ VIEWER", use_container_width=True, type="primary" if st.session_state.sys_menu == "viewer" else "secondary"):
        st.session_state.sys_menu = "viewer"
        st.rerun()
with m_col3:
    if st.button("⚙️ ADMINISTRATOR", use_container_width=True, type="primary" if st.session_state.sys_menu == "admin" else "secondary"):
        st.session_state.sys_menu = "admin"
        st.rerun()
with m_col4:
    if st.button("🔒 LOGOUT", use_container_width=True):
        st.session_state.main_authenticated = False
        st.session_state.sys_menu = "keyin"
        st.rerun()
        
st.markdown("<hr style='margin-top: 5px; margin-bottom: 20px; border-color: #cbd5e1;'>", unsafe_allow_html=True)

# ==========================================
# 💡 핵심 렌더링 함수: 대시보드 뷰어 통합 로직 (뷰어/어드민 공통 사용)
# ==========================================
def render_viewer(is_admin=False):
    # 💡 뷰어에만 적용되는 격리된 CSS (데이터 입력 폼과 충돌 방지)
    st.markdown("""
    <style>
    /* 라디오 버튼 우측 정렬 및 태블릿 줄바꿈 방지 */
    div[data-testid="stRadio"] { display: flex; justify-content: flex-end !important; width: 100%; margin-right: 0px !important; }
    div[role="radiogroup"] { justify-content: flex-end !important; flex-wrap: nowrap !important; gap: 15px !important; }
    div[role="radiogroup"] label { white-space: nowrap !important; }
    
    /* 카드 디자인 */
    .model-card { background: #000000; padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); flex: 1; }
    .model-card div, .model-card span { color: #FFFFFF !important; }
    .kpi-card { background: linear-gradient(135deg, #000000, #4472C4); padding: 20px; border-radius: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.1); flex: 1; }
    .kpi-card div, .kpi-card span { color: #FFC000 !important; }
    .metric-label { color: #1e293b !important; font-size: 1.1rem !important; font-weight: 800 !important; letter-spacing: 1px; margin-bottom: 10px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px; }
    .sbl-card { background: #ffffff; border: 1px solid #e2e8f0; border-left: 5px solid #ef4444; border-radius: 8px; padding: 12px; margin-bottom: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
    .sbl-title { color: #b91c1c !important; font-weight: 900; font-size: 0.95rem; margin-bottom: 5px; border-bottom: 1px solid #fecaca; padding-bottom: 3px; }
    .sbl-text { color: #334155 !important; font-size: 0.85rem; line-height: 1.5; font-weight: 500; }
    
    @keyframes blink { 0% { opacity: 1; box-shadow: 0 0 10px #3b82f6; } 50% { opacity: 0.3; box-shadow: 0 0 2px #3b82f6; } 100% { opacity: 1; box-shadow: 0 0 10px #3b82f6; } }
    .live-dot { height: 12px; width: 12px; background-color: #3b82f6; border-radius: 50%; display: inline-block; margin-right: 12px; margin-bottom: 2px; animation: blink 1.5s ease-in-out infinite; }
    </style>
    """, unsafe_allow_html=True)

    config = load_shared_config()
    if not config:
        st.warning("📡 설정값이 없습니다. ADMINISTRATOR 메뉴에서 설정을 완료하세요.")
        return

    if "viewer_time_range" not in st.session_state:
        st.session_state.viewer_time_range = config.get("time_range", "48H")

    # 💡 뷰어 전용 자바스크립트 자동화 스크립트 (어드민에서는 설정 입력 중 리셋 방지를 위해 비활성화)
    if not is_admin:
        auto_script = f"""
        <script>
        // 30분 자동 새로고침
        setTimeout(function() {{
            const btns = window.parent.document.querySelectorAll('button');
            for(let i=0; i<btns.length; i++){{
                if(btns[i].textContent && btns[i].textContent.includes('RELOAD')){{
                    btns[i].click();
                    break;
                }}
            }}
        }}, 1800000); 
        // 오토 로테이션
        {'setTimeout(function() { const btns = window.parent.document.querySelectorAll("button"); for(let i=0; i<btns.length; i++){ if(btns[i].textContent && btns[i].textContent.includes("Manual Rotate")){ btns[i].click(); break; } } }, 600000);' if config.get("auto_rotate_active", False) else ''}
        </script>
        """
        components.html(auto_script, height=0, width=0)

    # 💡 타이틀 및 컨트롤 버튼
    col1, col2 = st.columns([0.7, 0.3])
    with col1:
        mode_text = "ADMINISTRATOR VIEW (Preview Mode)" if is_admin else "Shared Dashboard (View Only)"
        st.markdown(f"<div class='command-header' style='font-size: 1.8rem; margin-top: 5px;'><span class='live-dot'></span>LIVE DASHBOARD SYSTEM</div>", unsafe_allow_html=True)
        st.markdown(f"<div style='color: #10b981; font-size: 0.85rem; margin-bottom: 15px; font-weight:bold;'>{mode_text}</div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        vc1, vc2 = st.columns(2)
        with vc1:
            if st.button("🔄 Manual Rotate", use_container_width=True, key=f"viewer_manual_rotate_{is_admin}"):
                st.session_state.rotate_idx += 1
                st.rerun()
        with vc2:
            if st.button("RELOAD", type="primary", use_container_width=True, key=f"viewer_reload_{is_admin}"):
                st.cache_data.clear()
                st.rerun()

    # 💡 라디오 버튼 (우측 끝 정렬)
    rad_c1, rad_c2 = st.columns([0.8, 0.2])
    with rad_c2:
        options_list = ["6H", "24H", "48H", "72H", "96H"]
        idx = options_list.index(st.session_state.viewer_time_range) if st.session_state.viewer_time_range in options_list else 2
        st.session_state.viewer_time_range = st.radio("조회 기간", options_list, index=idx, horizontal=True, label_visibility="collapsed", key=f'v_time_range_radio_{is_admin}')

    time_range = st.session_state.viewer_time_range
    now_kst = datetime.now(timezone(timedelta(hours=9))).replace(tzinfo=None)
    target_end_date = now_kst.date() 

    df = load_universal_data().copy()
    if df.empty: 
        st.warning("데이터베이스에 렌더링할 정보가 없습니다.")
        return

    def pct_to_float(x):
        try:
            if pd.isna(x) or str(x).strip() == '': return np.nan
            return float(str(x).replace('%', '').replace(',', '').strip())
        except: return np.nan
    def safe_int(x):
        try:
            if pd.isna(x) or str(x).strip() == '': return 0
            return int(float(str(x).replace(',', '').strip()))
        except: return 0
    def parse_lot(val):
        val_str = str(val).replace("'", "").strip()
        if val_str.endswith('.0'): val_str = val_str[:-2]
        if val_str.isdigit() and len(val_str) > 0: return val_str.zfill(5)
        return val_str if val_str else 'UNKNOWN'
        
    if 'LOT NO.' in df.columns: df['LOT NO.'] = df['LOT NO.'].apply(parse_lot)
    df['Yield_1'] = df.get('양품율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
    df['Yield_2'] = df.get('양품율(전/배 포함)', pd.Series([np.nan]*len(df))).apply(pct_to_float)
    df['검사수량'] = df.get('검사 수량', pd.Series([0]*len(df))).apply(safe_int)
    df['완전불량_Qty'] = df.get('완전불량', pd.Series([0]*len(df))).apply(safe_int)
    df['전면불량_Qty'] = df.get('전면불량', pd.Series([0]*len(df))).apply(safe_int)
    df['배면불량_Qty'] = df.get('배면불량', pd.Series([0]*len(df))).apply(safe_int)
    df['옵셋불량_Qty'] = df.get('옵셋불량', pd.Series([0]*len(df))).apply(safe_int)
    df['양품_Qty'] = df.get('양품수량', pd.Series([0]*len(df))).apply(safe_int)
    df['양품_FR_Qty'] = df.get('양품 수량(전/배 포함)', pd.Series([0]*len(df))).apply(safe_int)

    if df['Yield_1'].isna().all(): df['Yield_1'] = np.where(df['검사수량'] > 0, (df['양품_Qty'] / df['검사수량']) * 100, np.nan)
    if df['Yield_2'].isna().all(): df['Yield_2'] = np.where(df['검사수량'] > 0, (df['양품_FR_Qty'] / df['검사수량']) * 100, np.nan)
    df['Def_Comp'] = df.get('완전불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
    if df['Def_Comp'].isna().all(): df['Def_Comp'] = np.where(df['검사수량'] > 0, (df['완전불량_Qty'] / df['검사수량']) * 100, 0.0)
    df['Def_Front'] = df.get('전면불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
    if df['Def_Front'].isna().all(): df['Def_Front'] = np.where(df['검사수량'] > 0, (df['전면불량_Qty'] / df['검사수량']) * 100, 0.0)
    df['Def_Rear'] = df.get('배면불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
    if df['Def_Rear'].isna().all(): df['Def_Rear'] = np.where(df['검사수량'] > 0, (df['배면불량_Qty'] / df['검사수량']) * 100, 0.0)
    df['Def_Offset'] = df.get('옵셋불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
    if df['Def_Offset'].isna().all(): df['Def_Offset'] = np.where(df['검사수량'] > 0, (df['옵셋불량_Qty'] / df['검사수량']) * 100, 0.0)
    if '모델명(MI)' not in df.columns or df['모델명(MI)'].replace('', np.nan).isna().all(): df['모델명(MI)'] = 'ALL_MODELS'

    if time_range == "6H":
        target_start_dt = now_kst - timedelta(hours=6)
        df_target = df[df['DateTime'] >= target_start_dt].copy()
    else:
        if time_range == "24H": days_sub = 0
        elif time_range == "48H": days_sub = 1
        elif time_range == "72H": days_sub = 2
        else: days_sub = 3
        target_start_date = target_end_date - timedelta(days=days_sub)
        df_target = df[(df['DateOnly'] >= target_start_date) & (df['DateOnly'] <= target_end_date)].copy()

    display_std = config.get("sel_std", [])
    display_inc = config.get("sel_inc", [])
    model_color_dict = config.get("model_color_dict", {})
    sbl_limits = config.get("sbl_limits", {})
    all_selected = list(set(display_std + display_inc))

    if config.get("auto_rotate_active", False) and all_selected:
        current_idx = st.session_state.rotate_idx % len(all_selected)
        active_model = all_selected[current_idx]
        display_std = [active_model] if active_model in display_std else []
        display_inc = [active_model] if active_model in display_inc else []
        display_model_text = active_model
    else:
        if not all_selected: display_model_text = "ALL MODELS"
        elif len(all_selected) == 1: display_model_text = all_selected[0]
        else: display_model_text = ", ".join(all_selected[:2]) + ("..." if len(all_selected) > 2 else "")

    active_models_list = list(set(display_std + display_inc))
    base_df_active = df_target[df_target['모델명(MI)'].isin(active_models_list)].copy() if active_models_list else pd.DataFrame()

    def get_qty_metrics(df_sub):
        if df_sub.empty: return 0, 0, 0, 0, 0, 0
        t_ins = df_sub['검사수량'].sum()
        q_comp = df_sub['완전불량_Qty'].sum()
        q_front = df_sub['전면불량_Qty'].sum()
        q_rear = df_sub['배면불량_Qty'].sum()
        q_offset = df_sub['옵셋불량_Qty'].sum()
        q_good = 0
        for mod in display_std: q_good += df_sub[df_sub['모델명(MI)'] == mod]['양품_Qty'].sum()
        for mod in display_inc: q_good += df_sub[df_sub['모델명(MI)'] == mod]['양품_FR_Qty'].sum()
        return t_ins, q_good, q_comp, q_front, q_rear, q_offset

    yesterday_date = now_kst.date() - timedelta(days=1)
    o_t, o_g, o_c, o_f, o_r, o_o = get_qty_metrics(base_df_active)
    df_yesterday = base_df_active[base_df_active['DateOnly'] == yesterday_date].copy() if not base_df_active.empty else pd.DataFrame()
    y_t, y_g, y_c, y_f, y_r, y_o = get_qty_metrics(df_yesterday)
    df_6h = base_df_active[base_df_active['DateTime'] >= (now_kst - timedelta(hours=6))].copy() if not base_df_active.empty else pd.DataFrame()
    h_t, h_g, h_c, h_f, h_r, h_o = get_qty_metrics(df_6h)

    kpi_html = f"""
    <div style="display: flex; justify-content: space-between; gap: 15px; margin-bottom: 20px;">
        <div class="model-card">
            <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">적용 모델</div>
            <div style="font-size: 24px; font-weight: 900; margin-top: 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{display_model_text}</div>
        </div>
        <div class="kpi-card">
            <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">총 검사 수량</div>
            <div style="font-size: 28px; font-weight: 900; margin-top: 5px;">{o_t:,.0f} <span style="font-size: 14px; font-weight: normal;">EA</span></div>
        </div>
        <div class="kpi-card">
            <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">양품 수량</div>
            <div style="font-size: 28px; font-weight: 900; margin-top: 5px;">{o_g:,.0f} <span style="font-size: 14px; font-weight: normal;">EA</span></div>
        </div>
        <div class="kpi-card">
            <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">총 불량 수량</div>
            <div style="font-size: 28px; font-weight: 900; margin-top: 5px;">{o_c + o_f + o_r + o_o:,.0f} <span style="font-size: 14px; font-weight: normal;">EA</span></div>
        </div>
        <div class="kpi-card">
            <div style="font-size: 14px; font-weight: bold; opacity: 0.9;">종합 양품율</div>
            <div style="font-size: 28px; font-weight: 900; margin-top: 5px;">{(o_g/o_t*100) if o_t > 0 else 0:.1f} <span style="font-size: 14px; font-weight: normal;">%</span></div>
        </div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    col_left, col_mid, col_right = st.columns([0.22, 0.56, 0.22])

    with col_left:
        def make_donut_chart(title, t_ins, q_good, q_comp, q_front, q_rear, q_offset):
            labels = ['양품율', '완전불량', '전면불량', '배면불량', '옵셋불량']
            values = [q_good, q_comp, q_front, q_rear, q_offset]
            colors = ['#3B82F6', '#1E3A8A', '#FFC000', '#10B981', '#8B5CF6']
            l, v, c, txt = [], [], [], []
            for label, val, color in zip(labels, values, colors):
                if val > 0:
                    l.append(label)
                    v.append(val)
                    c.append(color)
                    pct = (val / t_ins * 100) if t_ins > 0 else 0
                    txt.append(f"{label}<br>{pct:.1f}%")
                    
            fig = go.Figure(data=[go.Pie(
                labels=l, values=v, hole=0.55,
                marker=dict(colors=c, line=dict(color='#ffffff', width=2)),
                textinfo='text', text=txt, textposition='outside', 
                textfont=dict(color='#0f172a', weight='bold', size=12, family="'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif"),
                hoverinfo='label+value', sort=False, direction='clockwise', rotation=270,
                domain=dict(x=[0.15, 0.85], y=[0.1, 0.9])
            )])
            fig.update_layout(
                title=dict(text=f"■ {title}", font=dict(color='#1e293b', size=14, weight='bold', family="'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif"), x=0.5, xanchor='center'),
                annotations=[dict(text=f"{t_ins:,.0f}<br><span style='font-size:11px; color:#64748b;'>Inspected</span>", 
                                  x=0.5, y=0.5, font_size=20, font_color='#1e293b', font_family="'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif", showarrow=False)],
                showlegend=False, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                margin=dict(l=40, r=40, t=50, b=30), height=260
            )
            return fig

        with st.container(border=True): st.plotly_chart(make_donut_chart(f"OVERALL ({time_range})", o_t, o_g, o_c, o_f, o_r, o_o), use_container_width=True, config={'displayModeBar': False}, theme=None)
        with st.container(border=True): st.plotly_chart(make_donut_chart("YESTERDAY", y_t, y_g, y_c, y_f, y_r, y_o), use_container_width=True, config={'displayModeBar': False}, theme=None)
        with st.container(border=True): st.plotly_chart(make_donut_chart("LAST 6 HOURS", h_t, h_g, h_c, h_f, h_r, h_o), use_container_width=True, config={'displayModeBar': False}, theme=None)

    with col_mid:
        if not base_df_active.empty:
            base_df_active['소요시간_num'] = pd.to_numeric(base_df_active['소요시간'].astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
            base_df_active = base_df_active.sort_values(['DateTime', '소요시간_num'], ascending=[True, True]).reset_index(drop=True)
            def clean_lot(val):
                val = str(val).replace("'", "").strip()
                if val.endswith('.0'): val = val[:-2]
                if val.isdigit() and len(val) > 0: return val.zfill(5)
                return val if val else 'UNKNOWN'
            base_df_active['LOT NO.'] = base_df_active.get('LOT NO.', pd.Series(['UNKNOWN']*len(base_df_active))).apply(clean_lot)
            base_df_active['HoverText'] = base_df_active.apply(lambda r: f"[{r.get('모델명(MI)', '')}]<br>Time: {r['DateTime'].strftime('%Y-%m-%d %H:%M')}<br>LOT: {r['LOT NO.']}", axis=1)

        # --- 2-1. YIELD TREND ---
        with st.container(border=True):
            fig_yld = go.Figure()
            y_min = 50.0
            if not base_df_active.empty:
                all_val = base_df_active['Yield_1'].dropna().tolist() + base_df_active['Yield_2'].dropna().tolist()
                if all_val: y_min = max(0, np.floor((min(all_val) - 5) / 10) * 10)
            if y_min > 80: y_min = 80.0

            if not base_df_active.empty:
                for mod in display_std:
                    m_df = base_df_active[base_df_active['모델명(MI)'] == mod].dropna(subset=['Yield_1'])
                    if m_df.empty: continue
                    c1 = model_color_dict.get(mod, '#3B82F6')
                    rgba_c1 = hex_to_rgba(c1, 0.15)
                    fig_yld.add_trace(go.Scatter(
                        x=m_df.index, y=m_df['Yield_1'], name=f"[{mod}] 양품율(기본)", 
                        mode='lines+markers', fill='tozeroy', fillcolor=rgba_c1,
                        line=dict(color=c1, width=3, shape='spline'), marker=dict(size=8, color=c1, symbol='circle'), hovertext=m_df['HoverText']
                    ))
                for mod in display_inc:
                    m_df = base_df_active[base_df_active['모델명(MI)'] == mod].dropna(subset=['Yield_2'])
                    if m_df.empty: continue
                    c1 = model_color_dict.get(mod, '#3B82F6')
                    rgba_c1 = hex_to_rgba(c1, 0.15)
                    fig_yld.add_trace(go.Scatter(
                        x=m_df.index, y=m_df['Yield_2'], name=f"[{mod}] 양품율(포함)", 
                        mode='lines+markers', fill='tozeroy', fillcolor=rgba_c1,
                        line=dict(color=c1, width=3, shape='spline'), marker=dict(size=8, color=c1, symbol='diamond'), hovertext=m_df['HoverText']
                    ))

            fig_yld.update_layout(
                title=dict(text=f"■ YIELD TREND ({time_range})", font=dict(color='#1e293b', size=16, weight='bold', family="'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif"), x=0.0, xanchor='left'),
                plot_bgcolor='#ffffff', paper_bgcolor='#ffffff',
                font=dict(color='#1e293b', family="'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif"),
                legend=dict(orientation="h", yanchor="bottom", y=1.15, xanchor="right", x=1, font=dict(color='#1e293b', size=12)), 
                margin=dict(l=60, r=30, t=80, b=60), height=380, hovermode='x unified'
            )
            
            if not base_df_active.empty:
                x_labels_yld = [r['DateTime'].strftime('%m-%d %H:%M') for _, r in base_df_active.iterrows()]
                fig_yld.update_xaxes(showgrid=True, gridcolor='#e2e8f0', linecolor='#94a3b8', tickmode='array', tickvals=base_df_active.index, ticktext=x_labels_yld, tickfont=dict(color='#1e293b', size=11))
            else:
                fig_yld.update_xaxes(showgrid=True, gridcolor='#e2e8f0', linecolor='#94a3b8', tickfont=dict(color='#1e293b', size=11))
                
            fig_yld.update_yaxes(title_text="양품율 (%)", range=[y_min, 100.0], tickformat=".1f", showgrid=True, gridcolor='#e2e8f0', linecolor='#94a3b8', tickfont=dict(color='#1e293b', size=11), title_font=dict(color='#1e293b', size=13), title_standoff=30)
            st.plotly_chart(fig_yld, use_container_width=True, config={'displayModeBar': False}, theme=None)
            
        # --- 2-2. DEFECT TREND ---
        with st.container(border=True):
            fig_def = go.Figure()
            
            if not base_df_active.empty:
                x_indices = base_df_active.index
                x_labels_def = [f"{r.get('도장일','')}<br>[{r.get('도장순서','')}]" for _, r in base_df_active.iterrows()]
                
                fig_def.add_trace(go.Bar(x=x_indices, y=base_df_active['Def_Front'], name='전면 불량율(%)', marker_color='#FFC000', text=base_df_active['Def_Front'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) and x>0 else ""), textposition='inside', textfont=dict(color='#000000', weight='bold'), hovertext=base_df_active['HoverText']))
                fig_def.add_trace(go.Bar(x=x_indices, y=base_df_active['Def_Rear'], name='배면 불량율(%)', marker_color='#10B981', text=base_df_active['Def_Rear'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) and x>0 else ""), textposition='inside', textfont=dict(color='#ffffff', weight='bold'), hovertext=base_df_active['HoverText']))
                fig_def.add_trace(go.Bar(x=x_indices, y=base_df_active['Def_Comp'], name='완전 불량율(%)', marker_color='#1E3A8A', text=base_df_active['Def_Comp'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) and x>0 else ""), textposition='inside', textfont=dict(color='#ffffff', weight='bold'), hovertext=base_df_active['HoverText']))
                fig_def.add_trace(go.Bar(x=x_indices, y=base_df_active['Def_Offset'], name='옵셋 불량율(%)', marker_color='#8B5CF6', text=base_df_active['Def_Offset'].apply(lambda x: f"{x:.1f}%" if pd.notna(x) and x>0 else ""), textposition='inside', textfont=dict(color='#ffffff', weight='bold'), hovertext=base_df_active['HoverText']))

            fig_def.update_layout(
                barmode='stack', bargap=0.2, 
                title=dict(text=f"■ DEFECT TREND ({time_range})", font=dict(color='#1e293b', size=16, weight='bold', family="'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif"), x=0.0, xanchor='left'),
                plot_bgcolor='#ffffff', paper_bgcolor='#ffffff',
                font=dict(color='#1e293b', family="'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif"),
                legend=dict(orientation="h", yanchor="bottom", y=1.15, xanchor="right", x=1, font=dict(color='#1e293b', size=12)),
                margin=dict(l=60, r=30, t=80, b=60), height=380, hovermode='x unified'
            )
            
            if not base_df_active.empty:
                fig_def.update_xaxes(title_text="도장일 [도장순서]", title_standoff=40, showgrid=False, linecolor='#94a3b8', tickmode='array', tickvals=x_indices, ticktext=x_labels_def, tickfont=dict(color='#1e293b', size=11), title_font=dict(color='#1e293b', size=13))
            else:
                fig_def.update_xaxes(title_text="도장일 [도장순서]", title_standoff=40, showgrid=False, linecolor='#94a3b8', tickfont=dict(color='#1e293b', size=11), title_font=dict(color='#1e293b', size=13))
                
            fig_def.update_yaxes(title_text="불량율 (%)", tickformat=".1f", showgrid=True, gridcolor='#e2e8f0', linecolor='#94a3b8', tickfont=dict(color='#1e293b', size=11), title_font=dict(color='#1e293b', size=13), title_standoff=30)
            st.plotly_chart(fig_def, use_container_width=True, config={'displayModeBar': False}, theme=None)

    with col_right:
        with st.container(border=True):
            st.markdown(f"<div class='metric-label' style='margin-top:5px; font-size:1.1rem;'>■ RECENT {time_range} ALERTS</div>", unsafe_allow_html=True)
            
            def render_sbl_list(d_col, title, is_yield=False):
                html = f"<div class='sbl-title' style='margin-top:10px;'>{title}</div>"
                if base_df_active.empty: return html + "<div class='sbl-text'>No data.</div>"
                
                sbl_items = []
                for _, r in base_df_active.iterrows():
                    mod = str(r.get('모델명(MI)', ''))
                    lot = str(r.get('LOT NO.', '')).replace("'", "")
                    t_str = r['DateTime'].strftime('%m-%d %H:%M')
                    
                    if is_yield:
                        val = r['Yield_2'] if mod in display_inc else r['Yield_1']
                        limit = sbl_limits.get(f'Yield_{mod}', sbl_limits.get('Yield_Default', 85.0))
                        if pd.notna(val) and 0 < val < limit:
                            sbl_items.append({'Time': t_str, 'Mod': mod, 'Lot': lot, 'Val': val, 'Limit': limit})
                    else:
                        val = r[d_col]
                        limit = sbl_limits.get(d_col, 5.0)
                        if pd.notna(val) and val > limit:
                            sbl_items.append({'Time': t_str, 'Mod': mod, 'Lot': lot, 'Val': val, 'Limit': limit})
                
                if not sbl_items:
                    html += "<div class='sbl-text' style='color:#94a3b8 !important; padding-bottom:5px;'>No alerts detected.</div>"
                else:
                    sbl_items = sorted(sbl_items, key=lambda x: x['Time'], reverse=True)[:5]
                    for item in sbl_items:
                        html += f"<div class='sbl-card'><div class='sbl-text'>[{item['Time']}] {item['Mod']}<br>LOT: {item['Lot']}<br><span style='color:#b91c1c; font-weight:bold;'>Value: {item['Val']:.1f}%</span> <span style='font-size:0.7rem; color:#64748b;'>(Limit: {item['Limit']:.1f}%)</span></div></div>"
                return html
            
            html_combined = f"""
            <div style='max-height: 720px; overflow-y: auto; padding-right: 5px; margin-bottom: 5px;'>
                {render_sbl_list('Yield_1', "Yield SBL List", is_yield=True)}
                <hr style='margin: 10px 0; border-color: #f1f5f9;'>
                {render_sbl_list('Def_Comp', "Complete Defect", is_yield=False)}
                <hr style='margin: 10px 0; border-color: #f1f5f9;'>
                {render_sbl_list('Def_Front', "Front Defect", is_yield=False)}
                <hr style='margin: 10px 0; border-color: #f1f5f9;'>
                {render_sbl_list('Def_Rear', "Rear Defect", is_yield=False)}
                <hr style='margin: 10px 0; border-color: #f1f5f9;'>
                {render_sbl_list('Def_Offset', "Offset Defect", is_yield=False)}
            </div>
            """
            st.markdown(html_combined, unsafe_allow_html=True)


# ==========================================
# 💡 [라우팅: 1. KEY-IN WIZARD 화면] (완벽한 구글 시트 양식 내장)
# ==========================================
if st.session_state.sys_menu == "keyin":
    st.markdown("<h3 style='color:#1e293b; font-weight:900;'>📝 VISION DATA KEY-IN WIZARD</h3>", unsafe_allow_html=True)
    st.info("💡 검사 데이터를 입력하면 양품율 및 불량율이 자동 계산되어 클라우드(Google Sheets)에 저장됩니다.")
    
    with st.form("keyin_form", clear_on_submit=False):
        st.markdown("<h4 style='color:#3b82f6;'>1. 기본 정보 입력</h4>", unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            input_date = st.date_input("날짜 (Date)")
            input_time = st.time_input("시작시간 (Time)")
        with col2:
            input_shift = st.selectbox("교대 (Shift)", ["주간", "야간"])
            input_worker = st.text_input("작업자 (Worker)")
        with col3:
            input_model = st.text_input("모델명(MI)")
            input_machine = st.text_input("호기 (Machine)")
        with col4:
            input_lot = st.text_input("LOT NO.")
            input_paint_date = st.text_input("도장일 (예: 9/23)")
            input_paint_seq = st.text_input("도장순서 (예: 1, 2, 3...)")
            
        st.markdown("<h4 style='color:#3b82f6; margin-top:20px;'>2. 검사 결과 입력 (수량)</h4>", unsafe_allow_html=True)
        col5, col6, col7 = st.columns(3)
        with col5:
            input_total = st.number_input("검사 수량 (Total)", min_value=0, step=1)
            input_good1 = st.number_input("양품 수량 (Yield 1)", min_value=0, step=1)
            input_good2 = st.number_input("양품 수량 (전/배 포함, Yield 2)", min_value=0, step=1)
        with col6:
            input_def_comp = st.number_input("완전 불량 (Complete Defect)", min_value=0, step=1)
            input_def_front = st.number_input("전면 불량 (Front Defect)", min_value=0, step=1)
        with col7:
            input_def_rear = st.number_input("배면 불량 (Rear Defect)", min_value=0, step=1)
            input_def_offset = st.number_input("옵셋 불량 (Offset Defect)", min_value=0, step=1)
            input_def_etc = st.number_input("기타 불량 (Etc)", min_value=0, step=1)

        input_note = st.text_input("비고 (Note)")
        st.markdown("<br>", unsafe_allow_html=True)
        submitted = st.form_submit_button("📤 데이터 구글 시트에 전송 및 저장", type="primary", use_container_width=True)

        if submitted:
            if not input_model or input_total <= 0:
                st.error("🚨 '모델명(MI)'과 '검사 수량'은 필수 입력 항목입니다. (수량은 1 이상이어야 합니다)")
            else:
                # 💡 데이터 자동 계산 로직
                def_total = input_def_comp + input_def_front + input_def_rear + input_def_offset + input_def_etc
                yield1 = (input_good1 / input_total * 100) if input_total > 0 else 0
                yield2 = (input_good2 / input_total * 100) if input_total > 0 else 0
                comp_rate = (input_def_comp / input_total * 100) if input_total > 0 else 0
                front_rate = (input_def_front / input_total * 100) if input_total > 0 else 0
                rear_rate = (input_def_rear / input_total * 100) if input_total > 0 else 0
                offset_rate = (input_def_offset / input_total * 100) if input_total > 0 else 0

                unique_id = str(uuid.uuid4())[:8].upper()

                row_data = []
                for col in EXCEL_COLUMNS:
                    if col == "고유 ID": row_data.append(unique_id)
                    elif col == "상태": row_data.append("완료")
                    elif col == "날짜": row_data.append(input_date.strftime("%Y-%m-%d"))
                    elif col == "교대": row_data.append(input_shift)
                    elif col == "시작시간": row_data.append(input_time.strftime("%H:%M:%S"))
                    elif col == "호기": row_data.append(input_machine)
                    elif col == "모델명(MI)": row_data.append(input_model)
                    elif col == "검사 수량": row_data.append(str(input_total))
                    elif col == "양품수량": row_data.append(str(input_good1))
                    elif col == "양품 수량(전/배 포함)": row_data.append(str(input_good2))
                    elif col == "불량수량": row_data.append(str(def_total))
                    elif col == "양품율": row_data.append(f"{yield1:.1f}%")
                    elif col == "양품율(전/배 포함)": row_data.append(f"{yield2:.1f}%")
                    elif col == "완전불량율": row_data.append(f"{comp_rate:.1f}%")
                    elif col == "전면불량율": row_data.append(f"{front_rate:.1f}%")
                    elif col == "배면불량율": row_data.append(f"{rear_rate:.1f}%")
                    elif col == "완전불량": row_data.append(str(input_def_comp))
                    elif col == "전면불량": row_data.append(str(input_def_front))
                    elif col == "배면불량": row_data.append(str(input_def_rear))
                    elif col == "옵셋불량": row_data.append(str(input_def_offset))
                    elif col == "기타": row_data.append(str(input_def_etc))
                    elif col == "도장일": row_data.append(input_paint_date)
                    elif col == "도장순서": row_data.append(input_paint_seq)
                    elif col == "LOT NO.": row_data.append(input_lot)
                    elif col == "작업자": row_data.append(input_worker)
                    elif col == "비고": row_data.append(input_note)
                    else: row_data.append("") 

                try:
                    with st.spinner('구글 시트에 데이터를 기록하는 중입니다...'):
                        doc = get_spreadsheet_doc()
                        ws = doc.worksheet(TAB_NAME)
                        ws.append_row(row_data)
                        st.success(f"✅ [{input_model}] 모델의 데이터가 구글 시트에 성공적으로 저장되었습니다!")
                        st.cache_data.clear() # 다음 뷰어 로드 시 즉시 반영되도록 캐시 클리어
                except Exception as e:
                    st.error(f"🚨 데이터 저장 실패: {e}")

# ==========================================
# 💡 [라우팅: 2. VIEWER 대시보드 화면]
# ==========================================
elif st.session_state.sys_menu == "viewer":
    render_viewer(is_admin=False)

# ==========================================
# 💡 [라우팅: 3. ADMINISTRATOR (관리자) 화면]
# ==========================================
elif st.session_state.sys_menu == "admin":
    st.markdown("<h3 style='color:#1e293b; font-weight:900;'>⚙️ ADMINISTRATOR CONTROL PANEL</h3>", unsafe_allow_html=True)
    
    config = load_shared_config() or {}
    df = load_universal_data()
    all_models = df['모델명(MI)'].dropna().unique().tolist() if not df.empty else ["ALL_MODELS"]
    
    with st.form("admin_config_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**1. 모니터링 대상 모델 선택**")
            sel_std = st.multiselect("기본 양품율 적용 모델 (Yield 1)", all_models, default=[m for m in config.get("sel_std", []) if m in all_models])
            sel_inc = st.multiselect("전/배 포함 양품율 적용 모델 (Yield 2)", all_models, default=[m for m in config.get("sel_inc", []) if m in all_models])
            
            st.markdown("<br>**2. 기본 설정**", unsafe_allow_html=True)
            time_range = st.selectbox("기본 조회 기간 (Default Time Range)", ["6H", "24H", "48H", "72H", "96H"], index=["6H", "24H", "48H", "72H", "96H"].index(config.get("time_range", "48H")))
            auto_rotate = st.checkbox("자동 로테이션 활성화 (10분 단위로 선택된 모델 순환 표출)", value=config.get("auto_rotate_active", False))
            
        with col2:
            st.markdown("**3. SBL (Sub-Block Limit) 알람 임계치 설정 (%)**")
            sbl_limits = config.get("sbl_limits", {})
            sbl_yield = st.number_input("📉 양품율 SBL (이하일 때 알람)", value=float(sbl_limits.get("Yield_Default", 85.0)), step=0.1)
            sbl_comp = st.number_input("📈 완전불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Comp", 5.0)), step=0.1)
            sbl_front = st.number_input("📈 전면불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Front", 5.0)), step=0.1)
            sbl_rear = st.number_input("📈 배면불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Rear", 5.0)), step=0.1)
            sbl_offset = st.number_input("📈 옵셋불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Offset", 5.0)), step=0.1)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.form_submit_button("💾 설정 저장 및 대시보드 클라우드 반영 (Save to Cloud)", use_container_width=True):
            new_config = {
                "sel_std": sel_std,
                "sel_inc": sel_inc,
                "time_range": time_range,
                "auto_rotate_active": auto_rotate,
                "sbl_limits": {
                    "Yield_Default": sbl_yield,
                    "Def_Comp": sbl_comp,
                    "Def_Front": sbl_front,
                    "Def_Rear": sbl_rear,
                    "Def_Offset": sbl_offset
                },
                "model_color_dict": config.get("model_color_dict", {})
            }
            if save_shared_config(new_config):
                st.cache_data.clear()
                st.rerun()

    st.markdown("<hr style='border-color: #cbd5e1; margin-top: 30px; margin-bottom: 30px;'>", unsafe_allow_html=True)
    
    # 설정이 반영된 뷰어 화면 렌더링
    render_viewer(is_admin=True)
