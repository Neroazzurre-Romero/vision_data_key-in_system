import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
import os
from datetime import datetime
import time
from io import BytesIO
from openpyxl.styles import Font
import openpyxl
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials

try:
    from PIL import Image
    import cv2
    from pyzbar.pyzbar import decode
    QR_AVAILABLE = True
except ImportError:
    QR_AVAILABLE = False

worker_a_list = ["A조", "작업자입력1", "작업자입력2"]
worker_b_list = ["B조", "작업자입력3", "작업자입력4"]
worker_c_list = ["C조", "작업자입력5", "작업자입력6"]
model_list = ["D65S(KRIOS)", "MEM", "Centaur", "Sphinx-E", "Banff", "AV-J", "Seattle", "Juliet-O"]

st.set_page_config(page_title="VISION DATA KEY-IN SYSTEM ----- (by. Romero)", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 영구 세션 상태 초기화 (입력 데이터 증발 방지)
# ==========================================
if "unlocked" not in st.session_state: st.session_state.unlocked = False
if "current_page" not in st.session_state: st.session_state.current_page = "input"
if "step" not in st.session_state: st.session_state.step = 1
if "unlocked" in st.query_params:
    st.session_state.unlocked = True
    st.query_params.clear()

default_state = {
    "work_date": datetime.now().date(), "shift_type": "주간", 
    "worker_a": "A조", "worker_b": "B조", "worker_c": "C조",
    "model_name": "D65S(KRIOS)", "lot_input_field": "", "in_date_field": datetime.now().date(),
    "plating_type": "A", "start_date": datetime.now().date(), "start_time": datetime.now().time(),
    "end_date": datetime.now().date(), "end_time": datetime.now().time(), "unit": "1호기",
    "category": "1차 검사", "idle_time": 0, "painting_date": datetime.now().date(),
    "painting_order": 1, "painting_line": "B Line", "clip_val": "1",
    "base_val": "1", "cover_val": "1", "assembler_val": "선택안함",
    "good_qty": 0, "comp_def": 0, "front_def": 0, "rear_def": 0, "offset_def": 0,
    "shortage_qty": 0, "etc_def": 0, "oqc_status": "선택안함", "remarks": "",
    "scanned_raw_data": "", "comp_warned": False, "front_warned": False, 
    "rear_warned": False, "offset_warned": False
}

for key, value in default_state.items():
    if key not in st.session_state:
        st.session_state[key] = value

# ==========================================
# 메인 잠금 화면 (Splash Screen)
# ==========================================
if not st.session_state.unlocked:
    hide_sidebar_style = """
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        #MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;}
    </style>
    """
    st.markdown(hide_sidebar_style, unsafe_allow_html=True)
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        st.markdown("<h1 style='text-align: center; color: #1e293b; font-size: 45px; font-weight: 900;'>VISION DATA KEY-IN SYSTEM</h1>", unsafe_allow_html=True)
        st.markdown("<br><br>", unsafe_allow_html=True)
        
        if st.button("UNLOCK_SYSTEM_BTN_HIDDEN"):
            st.session_state.unlocked = True
            st.rerun()

        slider_html = """
        <div id="slider-container" style="background: #ffffff; border: 2px solid #e2e8f0; border-radius: 40px; position: relative; width: 100%; max-width: 400px; height: 68px; margin: 0 auto; overflow: hidden; display: flex; align-items: center; box-shadow: inset 0 2px 5px rgba(0,0,0,0.05);">
            <div id="slider-fill" style="position: absolute; left: 0; top: 0; height: 100%; width: 0; background-color: #3b82f6; border-radius: 40px 0 0 40px;"></div>
            <div id="slider-text" style="position: absolute; width: 100%; text-align: center; color: #94a3b8; font-size: 20px; font-weight: bold; font-family: sans-serif; pointer-events: none; z-index: 2; transition: color 0.3s;">Slide to Unlock</div>
            <div id="slider-thumb" style="position: absolute; left: 4px; width: 56px; height: 56px; background: #ffffff; border-radius: 50%; box-shadow: 0 2px 6px rgba(0,0,0,0.2); cursor: pointer; z-index: 3; display: flex; align-items: center; justify-content: center; color: #3b82f6; font-size: 24px;">▶</div>
        </div>
        <script>
            const container = document.getElementById('slider-container');
            const thumb = document.getElementById('slider-thumb');
            const fill = document.getElementById('slider-fill');
            const text = document.getElementById('slider-text');

            const unlockSystem = () => {
                const btns = window.parent.document.querySelectorAll('button');
                for(let b of btns) { if(b.innerText.includes('UNLOCK_SYSTEM_BTN_HIDDEN')) { b.click(); break; } }
            };

            const btns = window.parent.document.querySelectorAll('button');
            for(let b of btns) { if(b.innerText.includes('UNLOCK_SYSTEM_BTN_HIDDEN')) { b.style.display = 'none'; } }

            let isDragging = false;
            let startX, currentX = 0;

            function startDrag(e) {
                isDragging = true;
                let clientX = e.type.includes('mouse') ? e.clientX : e.touches[0].clientX;
                startX = clientX - currentX;
            }

            function drag(e) {
                if (!isDragging) return;
                if(e.cancelable) e.preventDefault();
                let clientX = e.type.includes('mouse') ? e.clientX : e.touches[0].clientX;
                currentX = clientX - startX;
                const maxDrag = container.clientWidth - thumb.clientWidth - 8; 
                if (currentX < 0) currentX = 0;
                if (currentX > maxDrag) currentX = maxDrag;
                thumb.style.transform = `translateX(${currentX}px)`;
                fill.style.width = (currentX + thumb.clientWidth / 2) + 'px';
                if (currentX > maxDrag * 0.4) { text.style.color = '#ffffff'; } else { text.style.color = '#94a3b8'; }
                if (currentX >= maxDrag) {
                    isDragging = false;
                    text.innerText = "Unlocked!";
                    thumb.innerHTML = "✔";
                    setTimeout(() => { unlockSystem(); }, 200);
                }
            }

            function endDrag(e) {
                if (!isDragging) return;
                isDragging = false;
                const maxDrag = container.clientWidth - thumb.clientWidth - 8;
                if (currentX < maxDrag) {
                    thumb.style.transition = 'transform 0.3s ease';
                    fill.style.transition = 'width 0.3s ease';
                    currentX = 0;
                    thumb.style.transform = `translateX(0px)`;
                    fill.style.width = '0px';
                    text.style.color = '#94a3b8';
                    setTimeout(() => { thumb.style.transition = 'none'; fill.style.transition = 'none'; }, 300);
                }
            }

            thumb.addEventListener('mousedown', startDrag); document.addEventListener('mousemove', drag); document.addEventListener('mouseup', endDrag);
            thumb.addEventListener('touchstart', startDrag, {passive: false}); document.addEventListener('touchmove', drag, {passive: false}); document.addEventListener('touchend', endDrag);
        </script>
        """
        components.html(slider_html, height=90)
            
    st.markdown("<div style='position: fixed; bottom: 10%; left: 0; width: 100%; text-align: center; font-size: 10pt; color: #94a3b8; font-weight: bold;'>vision data key-in system --- Romero.K</div>", unsafe_allow_html=True)
    st.stop()

# ----------------------------------------------------
# 마법 코드 1: UI 디자인 커스텀 및 사이드바 스타일링
# ----------------------------------------------------
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;} footer {visibility: hidden;} header {visibility: hidden;} body { overscroll-behavior-y: none !important; } ::-webkit-scrollbar { display: none; }
.block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; }
div[data-testid="stMarkdownContainer"] p strong, div[data-testid="stWidgetLabel"] p, div[data-testid="stWidgetLabel"] p strong { font-size: 1.15rem !important; font-weight: 800 !important; color: #1e293b !important; }
div[data-baseweb="input"] > div, div[data-baseweb="select"] > div { min-height: 3.5rem !important; }
div[data-baseweb="input"] input, div[data-baseweb="select"] div { font-size: 1.15rem !important; }
div[data-baseweb="textarea"] textarea { font-size: 1.15rem !important; min-height: 100px !important; }
button[kind="primary"] { background-color: #4b6584 !important; color: white !important; border: none !important; font-size: 16px !important; font-weight: bold !important; padding: 10px !important; }
button[kind="primary"]:hover { background-color: #3b5068 !important; }
[data-testid="stSidebar"] { background: linear-gradient(135deg, #0f172a 0%, #020617 100%) !important; }
[data-testid="stSidebar"] * { color: #f8fafc !important; }
[data-testid="stSidebar"] .stButton > button { height: 65px !important; justify-content: flex-start !important; padding-left: 15px !important; margin-bottom: 6px !important; border-radius: 8px !important; background-color: transparent !important; color: #f8fafc !important; border: 1px solid #334155 !important; }
[data-testid="stSidebar"] .stButton > button p { font-weight: 800 !important; font-size: 16px !important; text-indent: 10px !important; text-align: left !important; }
[data-testid="stSidebar"] .stButton > button[kind="primary"] { background-color: #3b82f6 !important; color: white !important; border: 1px solid #2563eb !important; }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

components.html(
    """
    <script>
    if (window.parent && !window.parent.appPluginLoadedFull) {
        window.parent.appPluginLoadedFull = true;
        
        const formatNavButtons = () => {
            if (!window.parent.document) return;
            const buttons = window.parent.document.querySelectorAll('button');
            buttons.forEach(btn => {
                const text = btn.innerText || "";
                if (text.includes('⬅️ 이전')) { btn.style.backgroundColor = '#FFC000'; btn.style.color = '#000000'; btn.style.border = 'none'; }
                if (text.includes('다음 ➡️')) { btn.style.backgroundColor = '#00B050'; btn.style.color = '#FFFFFF'; btn.style.border = 'none'; }
                
                if (text.includes('데이터 최종 저장')) { 
                    btn.style.height = '100px'; 
                    btn.style.marginTop = '28px'; 
                    btn.style.fontSize = '18px'; 
                    btn.style.whiteSpace = 'pre-wrap'; 
                }
                if (text.trim() === 'Data Analysis') { 
                    btn.style.height = '58px'; 
                    btn.style.fontSize = '16px'; 
                    btn.style.marginTop = '0px'; 
                }
            });
        };
        const styleScanner = () => {
            if (!window.parent.document) return;
            const targets = window.parent.document.querySelectorAll('div[id="scanner_target"]');
            targets.forEach(t => {
                let parent = t.parentElement;
                while(parent && parent.getAttribute('data-testid') !== 'stVerticalBlock') { parent = parent.parentElement; }
                if(parent && !parent.dataset.styled) {
                    parent.style.backgroundColor = '#D9E1F2';
                    parent.style.padding = '20px';
                    parent.style.borderRadius = '12px';
                    parent.style.boxShadow = '0 4px 6px rgba(0,0,0,0.1)';
                    parent.style.marginBottom = '20px';
                    parent.dataset.styled = 'true';
                }
            });
        };
        const disableKeyboard = () => {
            if (!window.parent.document) return;
            window.parent.document.querySelectorAll('input').forEach(el => {
                const placeholder = el.getAttribute('placeholder') || '';
                if (placeholder.includes('YYYY') || placeholder.includes('HH:MM')) {
                    if (el.getAttribute('inputmode') !== 'none') el.setAttribute('inputmode', 'none');
                }
            });
        };
        const observer = new MutationObserver(() => { disableKeyboard(); formatNavButtons(); styleScanner(); });
        if (window.parent.document.body) { observer.observe(window.parent.document.body, { childList: true, subtree: true }); }
        disableKeyboard(); formatNavButtons(); styleScanner();
    }
    </script>
    """, height=0, width=0
)

# ----------------------------------------------------
# 터치형 박스(Grid) 렌더링 함수
# ----------------------------------------------------
def render_grid_buttons(options, state_key, columns):
    rows = [options[i:i+columns] for i in range(0, len(options), columns)]
    for row_opts in rows:
        cols = st.columns(columns)
        for i, opt in enumerate(row_opts):
            with cols[i]:
                if opt.strip() == "": st.write("") 
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

@st.cache_resource(ttl=600)
def get_sheet():
    for attempt in range(3):
        try:
            creds_data = st.secrets["google_credentials"]
            clean_data = creds_data.strip().strip("'").strip('"') if isinstance(creds_data, str) else dict(creds_data)
            creds_dict = json.loads(clean_data, strict=False) if isinstance(creds_data, str) else clean_data
            if "private_key" in creds_dict: creds_dict["private_key"] = creds_dict["private_key"].replace('\\n', '\n')
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPE)
            doc = gspread.authorize(creds).open_by_key(SPREADSHEET_ID)
            try: return doc.worksheet(TAB_NAME)
            except: return doc.sheet1
        except Exception as e:
            if "503" in str(e) and attempt < 2:
                time.sleep(2)
                continue
            return None

@st.cache_data(ttl=60)
def load_data():
    sheet = get_sheet()
    if sheet is None: return pd.DataFrame(columns=EXCEL_COLUMNS)
    try:
        raw_data = sheet.get_all_values()
        valid_data = [row for row in raw_data if any(str(cell).strip() for cell in row)]
        if len(valid_data) < 2: return pd.DataFrame(columns=EXCEL_COLUMNS)
        header_idx = 0
        for i, row in enumerate(valid_data[:10]):
            row_str = "".join(str(c).replace(" ", "") for c in row)
            if "날짜" in row_str or "교대" in row_str or "모델명" in row_str:
                header_idx = i; break
        headers = [str(h).strip() for h in valid_data[header_idx]]
        df = pd.DataFrame(valid_data[header_idx+1:])
        df.columns = headers[:len(df.columns)]
        clean_headers = {c.replace(" ", "").upper(): c for c in df.columns}
        result_df = pd.DataFrame(index=df.index)
        for col in EXCEL_COLUMNS:
            col_key = col.replace(" ", "").upper()
            if col_key in clean_headers: result_df[col] = df[clean_headers[col_key]]
            else: result_df[col] = "" 
        return result_df
    except: return pd.DataFrame(columns=EXCEL_COLUMNS)

def save_data_append(df):
    sheet = get_sheet()
    if sheet is None: return False
    try:
        header_check = sheet.row_values(1)
        if not header_check: sheet.append_row(EXCEL_COLUMNS, value_input_option='USER_ENTERED')
        records = []
        for _, row in df.iterrows():
            records.append(["" if str(row.get(col, "")).strip().lower() in ["nan", "none"] else str(row.get(col, "")).strip() for col in EXCEL_COLUMNS])
        sheet.append_rows(records, value_input_option='USER_ENTERED')
        load_data.clear() 
        return True
    except Exception as e:
        st.error(f"데이터 저장 오류: {e}")
        return False

# ----------------------------------------------------
# 팝업(모달) SBL 함수
# ----------------------------------------------------
@st.dialog("SBL Warning!")
def show_sbl_warning(defect_type, rate):
    st.markdown(f"### [{defect_type}] 불량 제품 별도 보관 조치")
    st.error(f"현재 1차검사 공정의 {defect_type}율이 **{rate:.1f}%** 로 기준치(5.0%)를 초과하였습니다.")
    if st.button("확인 완료 (닫기)", key=f"btn_close_{defect_type}"):
        st.rerun()

# ==========================================
# 메인 프로세스 화면 구성
# ==========================================
if st.session_state.current_page == "analysis":
    st.markdown("## 종합 생산 데이터 분석")
    if st.button("뒤로 가기 (데이터 입력 화면으로)"):
        st.session_state.current_page = "input"
        st.rerun()
    df = load_data().copy()
    if df.empty: st.warning("데이터가 없습니다.")
    else: st.dataframe(df.head(50))

elif st.session_state.current_page == "input":
    
    # 💡 비율 8.5:1.5 및 연한 회색 배경 스타일, 버튼 높이 동기화
    top_c1, top_c2 = st.columns([0.85, 0.15])
    with top_c1:
        st.markdown("""
            <div style='background: #e2e8f0; padding: 0 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 2px 4px rgba(0,0,0,0.05); border: 1px solid #cbd5e1; height: 58px; display: flex; align-items: center;'>
                <h3 style='color: #1e293b; margin: 0; font-weight: 800;'>VISION DATA KEY-IN SYSTEM</h3>
            </div>
        """, unsafe_allow_html=True)
    with top_c2:
        if st.button("Data Analysis", use_container_width=True, type="primary"):
            st.session_state.current_page = "analysis"
            st.rerun()

    with st.sidebar:
        steps_titles = [
            "생산 등록", "작업 정보", "Coating Data", "Assemble Data", 
            "VISION Data", "Report & History"
        ]
        for i, title in enumerate(steps_titles, 1):
            btn_type = "primary" if st.session_state.step == i else "secondary"
            if st.button(title, key=f"nav_btn_{i}", type=btn_type, use_container_width=True):
                st.session_state.step = i
                st.rerun()

        st.markdown("<br><hr style='border-color: #334155; margin-top: 5px; margin-bottom: 15px;'>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.session_state.step > 1:
                if st.button("⬅️ 이전", use_container_width=True):
                    st.session_state.step -= 1
                    st.rerun()
        with c2:
            if st.session_state.step < 6:
                if st.button("다음 ➡️", use_container_width=True):
                    st.session_state.step += 1
                    st.rerun()

        # 💡 로고 대체: create by 텍스트
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div style='text-align: center; color: #94a3b8; font-size: 13px; font-weight: bold;'>create by --- Romero.K</div>", unsafe_allow_html=True)

    step = st.session_state.step

    def parse_scanned_data():
        raw_val = st.session_state.scanned_raw_data
        if not raw_val: return
        if '$' in raw_val:
            parts = [p for p in raw_val.split('$') if p]
            if len(parts) >= 5:
                plating_code = parts[2]
                if plating_code == 'S110': st.session_state.plating_type = 'A'
                elif plating_code == 'S112': st.session_state.plating_type = 'B'
                
                date_str = parts[3]
                if len(date_str) == 8 and date_str.isdigit():
                    try: st.session_state.in_date_field = datetime.strptime(date_str, "%Y%m%d").date()
                    except ValueError: pass
                
                st.session_state.lot_input_field = parts[4]
            else:
                st.session_state.lot_input_field = parts[-1]
        else:
            st.session_state.lot_input_field = raw_val

    if step == 1:
        c1, c2 = st.columns(2)
        with c1: st.session_state.work_date = st.date_input("**근무일자**", value=st.session_state.work_date)
        with c2: st.session_state.model_name = st.selectbox("**모델명**", model_list, index=model_list.index(st.session_state.model_name) if st.session_state.model_name in model_list else 0)
        
        st.markdown("**교대**")
        render_grid_buttons(["주간", "야간"], "shift_type", 2)
        
        st.markdown("**작업자**")
        w_col1, w_col2, w_col3 = st.columns(3)
        with w_col1: st.session_state.worker_a = st.selectbox("A조", worker_a_list, index=worker_a_list.index(st.session_state.worker_a) if st.session_state.worker_a in worker_a_list else 0, label_visibility="collapsed")
        with w_col2: st.session_state.worker_b = st.selectbox("B조", worker_b_list, index=worker_b_list.index(st.session_state.worker_b) if st.session_state.worker_b in worker_b_list else 0, label_visibility="collapsed")
        with w_col3: st.session_state.worker_c = st.selectbox("C조", worker_c_list, index=worker_c_list.index(st.session_state.worker_c) if st.session_state.worker_c in worker_c_list else 0, label_visibility="collapsed")

        st.markdown("<hr>", unsafe_allow_html=True)
        with st.container():
            st.markdown("<div id='scanner_target'></div>", unsafe_allow_html=True)
            
            sc1, sc2, sc3 = st.columns([1, 1.5, 1])
            with sc1:
                if st.button("📷 QR CODE SCANNER", use_container_width=True, type="primary"):
                    st.info("👉 우측 입력창을 터치하여 태블릿에 설치된 '스캐너 키보드 앱'을 실행하세요.")
            with sc2:
                st.text_input("스캔 데이터", key="scanned_raw_data", on_change=parse_scanned_data, label_visibility="collapsed", placeholder="터치하여 스캐너 앱 실행")
            with sc3:
                if st.button("적용", type="primary", use_container_width=True):
                    parse_scanned_data()
                    st.rerun()
                
        st.markdown("<br>", unsafe_allow_html=True)
        c_res1, c_res2 = st.columns(2)
        with c_res1:
            st.text_input("**LOT (적용됨)**", value=st.session_state.lot_input_field, disabled=True)
        with c_res2:
            st.date_input("**입고일 (적용됨)**", value=st.session_state.in_date_field, disabled=True)
            
        st.markdown("**도금 구분 (적용됨)**")
        render_grid_buttons(["A", "B"], "plating_type", 2)

    elif step == 2:
        c1, c2 = st.columns(2)
        with c1: st.session_state.start_date = st.date_input("**시작일**", value=st.session_state.start_date)
        with c2: st.session_state.start_time = st.time_input("**시작시간**", value=st.session_state.start_time)
        
        c3, c4 = st.columns(2)
        with c3: st.session_state.end_date = st.date_input("**종료일**", value=st.session_state.end_date)
        with c4: st.session_state.end_time = st.time_input("**종료시간**", value=st.session_state.end_time)
        
        st.markdown("**호기**")
        render_grid_buttons(["1호기", "2호기", "3호기", " "], "unit", 2)
        
        st.markdown("**검사 구분**")
        render_grid_buttons(["1차 검사", "2차 검사"], "category", 2)
        
        st.session_state.idle_time = st.number_input("**휴동시간 (분)**", min_value=0, value=st.session_state.idle_time)
        
        start_dt = datetime.combine(st.session_state.start_date, st.session_state.start_time)
        end_dt = datetime.combine(st.session_state.end_date, st.session_state.end_time)
        raw_duration = int((end_dt - start_dt).total_seconds() / 60)
        duration_minutes = max(0, raw_duration - st.session_state.idle_time)
        st.text_input("**소요시간 (휴동시간 차감됨)**", value=f"{duration_minutes:,} 분", disabled=True)

    elif step == 3:
        c1, c2 = st.columns(2)
        with c1: st.session_state.painting_date = st.date_input("**도장일**", value=st.session_state.painting_date)
        with c2: st.session_state.painting_order = st.number_input("**도장순서**", min_value=1, value=st.session_state.painting_order)
        
        st.markdown("**도장라인**")
        render_grid_buttons(["A Line", "B Line", "C Line"], "painting_line", 3)

    elif step == 4:
        num_options = ["1"] + [str(i) for i in range(2, 11)] + ["선택안함"]
        
        c1, c2, c3 = st.columns(3)
        with c1: st.session_state.clip_val = st.selectbox("**CLIP**", num_options, index=num_options.index(st.session_state.clip_val) if st.session_state.clip_val in num_options else 0)
        with c2: st.session_state.base_val = st.selectbox("**BASE**", num_options, index=num_options.index(st.session_state.base_val) if st.session_state.base_val in num_options else 0)
        with c3: st.session_state.cover_val = st.selectbox("**COVER**", num_options, index=num_options.index(st.session_state.cover_val) if st.session_state.cover_val in num_options else 0)

        st.markdown("**조립기**")
        render_grid_buttons(["1호기", "2호기", "3호기", "4호기"], "assembler_val", 2)

    elif step == 5:
        q1, q2, q3 = st.columns(3)
        with q2: 
            st.session_state.good_qty = st.number_input("**양품수량**", min_value=0, value=st.session_state.good_qty)
        
        st.markdown("**🚨 불량 세부**")
        c1, c2, c3 = st.columns(3)
        with c1: st.session_state.comp_def = st.number_input("**완전불량**", min_value=0, value=st.session_state.comp_def)
        with c2: st.session_state.front_def = st.number_input("**전면불량**", min_value=0, value=st.session_state.front_def)
        with c3: st.session_state.rear_def = st.number_input("**배면불량**", min_value=0, value=st.session_state.rear_def)
        
        c4, c5, c6, c7 = st.columns(4)
        with c4: st.session_state.offset_def = st.number_input("**옵셋불량**", min_value=0, value=st.session_state.offset_def)
        with c5: st.session_state.shortage_qty = st.number_input("**수량부족**", min_value=0, value=st.session_state.shortage_qty)
        with c6: st.session_state.etc_def = st.number_input("**기타**", min_value=0, value=st.session_state.etc_def)
        with c7: st.session_state.oqc_status = st.selectbox("**OQC**", ["선택안함", "육안", "OQC"], index=["선택안함", "육안", "OQC"].index(st.session_state.oqc_status))
        
        bad_qty = st.session_state.comp_def + st.session_state.front_def + st.session_state.rear_def + st.session_state.offset_def + st.session_state.etc_def
        total_qty = max(0, st.session_state.good_qty + bad_qty - st.session_state.shortage_qty)
        
        with q1: 
            st.text_input("**검사 수량 (자동)**", value=f"{total_qty:,}", disabled=True)
        with q3: 
            st.text_input("**불량수량 (자동)**", value=f"{bad_qty:,}", disabled=True)

        if st.session_state.category == "1차 검사" and total_qty > 0:
            comp_rate = (st.session_state.comp_def / total_qty) * 100
            front_rate = (st.session_state.front_def / total_qty) * 100
            rear_rate = (st.session_state.rear_def / total_qty) * 100
            offset_rate = (st.session_state.offset_def / total_qty) * 100
            
            if comp_rate > 5.0 and not st.session_state.comp_warned:
                show_sbl_warning("완전불량", comp_rate)
                st.session_state.comp_warned = True
            if front_rate > 5.0 and not st.session_state.front_warned:
                show_sbl_warning("전면불량", front_rate)
                st.session_state.front_warned = True
            if rear_rate > 5.0 and not st.session_state.rear_warned:
                show_sbl_warning("배면불량", rear_rate)
                st.session_state.rear_warned = True
            if offset_rate > 5.0 and not st.session_state.offset_warned:
                show_sbl_warning("옵셋불량", offset_rate)
                st.session_state.offset_warned = True

        st.markdown("<hr><b>📈 수율 현황</b>", unsafe_allow_html=True)
        rate_good = round((st.session_state.good_qty / total_qty) * 100, 1) if total_qty > 0 else 0.0
        
        c_yield, c_comp, c_front, c_rear, c_offset = "#002b5e", "#b22222", "#ed7d31", "#00b050", "#7030a0"
        
        fig_donut = go.Figure(go.Pie(
            labels=['양품율', '불량율'], values=[rate_good, 100-rate_good if rate_good > 0 else 0], 
            hole=.65, sort=False, direction='clockwise',
            marker=dict(colors=[c_yield, '#e2e8f0'], line=dict(color='#ffffff', width=2)), 
            hoverinfo="label+percent", textinfo="none"
        ))
        fig_donut.update_layout(
            showlegend=False, height=250, margin=dict(t=10, b=10, l=10, r=10),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            annotations=[dict(text=f"{rate_good:.1f}%", x=0.5, y=0.5, font_size=30, font_color=c_yield, showarrow=False)]
        )
        
        df_defects = pd.DataFrame({
            "불량 항목": ['완전불량', '전면불량', '배면불량', '옵셋불량'], 
            "비율 (%)": [
                round((st.session_state.comp_def/total_qty)*100,1) if total_qty>0 else 0,
                round((st.session_state.front_def/total_qty)*100,1) if total_qty>0 else 0,
                round((st.session_state.rear_def/total_qty)*100,1) if total_qty>0 else 0,
                round((st.session_state.offset_def/total_qty)*100,1) if total_qty>0 else 0
            ]
        })
        y_max = max(df_defects["비율 (%)"]) * 1.4 if not df_defects.empty and max(df_defects["비율 (%)"]) > 0 else 5
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(x=df_defects["불량 항목"], y=[y_max]*4, marker_color='#f1f5f9', hoverinfo='none', width=0.45))
        fig_bar.add_trace(go.Bar(x=df_defects["불량 항목"], y=df_defects["비율 (%)"], marker_color=[c_comp, c_front, c_rear, c_offset], width=0.45, texttemplate=''))
        fig_bar.add_trace(go.Scatter(
            x=df_defects["불량 항목"], y=df_defects["비율 (%)"], mode='markers+text',
            marker=dict(size=40, color=[c_comp, c_front, c_rear, c_offset], line=dict(color='white', width=3)),
            text=df_defects["비율 (%)"].apply(lambda x: f"{x:.1f}"), textfont=dict(color='white', size=14, weight='bold'),
            textposition='middle center', hoverinfo='none'
        ))
        fig_bar.update_layout(barmode='overlay', showlegend=False, height=250, margin=dict(t=10, b=20, l=10, r=10), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(showgrid=False), yaxis=dict(showgrid=False, showticklabels=False, range=[0, y_max]))

        g_col1, g_col2 = st.columns(2)
        with g_col1: st.plotly_chart(fig_donut, use_container_width=True)
        with g_col2: st.plotly_chart(fig_bar, use_container_width=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        
        rem_col, save_col = st.columns([0.7, 0.3])
        with rem_col:
            st.session_state.remarks = st.text_area("**비고**", value=st.session_state.remarks, height=100)
            
        with save_col:
            if st.button("💾 데이터 최종 저장\n(구글 시트 전송)", type="primary", use_container_width=True):
                if total_qty == 0: st.warning("입력된 데이터(검사수량)가 없습니다.")
                elif not st.session_state.lot_input_field: st.warning("LOT 번호를 1단계에서 확인해주세요.")
                else:
                    with st.spinner("저장 중..."):
                        start_dt = datetime.combine(st.session_state.start_date, st.session_state.start_time)
                        end_dt = datetime.combine(st.session_state.end_date, st.session_state.end_time)
                        raw_duration = int((end_dt - start_dt).total_seconds() / 60)
                        duration_minutes = max(0, raw_duration - st.session_state.idle_time)

                        uph_val = int((total_qty / duration_minutes) * 60) if duration_minutes > 0 else 0
                        upd_val = uph_val * 22
                        good_include_front_rear = st.session_state.good_qty + st.session_state.front_def + st.session_state.rear_def
                        
                        if total_qty > 0:
                            rate_good = round((st.session_state.good_qty / total_qty) * 100, 1)
                            rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
                            comp_rate_num = round(st.session_state.comp_def / total_qty * 100, 1)
                            front_rate_num = round(st.session_state.front_def / total_qty * 100, 1)
                            rear_rate_num = round(st.session_state.rear_def / total_qty * 100, 1)
                            offset_rate_num = round(st.session_state.offset_def / total_qty * 100, 1)
                        else:
                            rate_good = rate_good_inc = comp_rate_num = front_rate_num = rear_rate_num = offset_rate_num = 0.0

                        fmt_date = f"{st.session_state.work_date.month}/{st.session_state.work_date.day}"
                        fmt_paint_date = f"{st.session_state.painting_date.month}/{st.session_state.painting_date.day}" 
                        fmt_in_date = st.session_state.in_date_field.strftime("%Y-%m-%d") 
                        fmt_paint_line = st.session_state.painting_line.replace(" Line", "") if st.session_state.painting_line != "선택안함" else ""
                        fmt_assembler = st.session_state.assembler_val.replace("호기", "") if st.session_state.assembler_val != "선택안함" else ""
                        
                        workers = [w for w in [st.session_state.worker_a, st.session_state.worker_b, st.session_state.worker_c] if w not in ["A조", "B조", "C조"]]
                        fmt_worker = ", ".join(workers) if workers else ""

                        new_data = pd.DataFrame([{
                            "날짜": fmt_date, "교대": st.session_state.shift_type,
                            "시작시간": start_dt.strftime("%H:%M"), "종료시간": end_dt.strftime("%H:%M"),
                            "휴동시간": f"{st.session_state.idle_time:,}", "소요시간": f"{duration_minutes:,}", "구분": st.session_state.category, "호기": st.session_state.unit, 
                            "모델명(MI)": st.session_state.model_name, "도금구분": st.session_state.plating_type, "UPH": f"{uph_val:,}", "UPD": f"{upd_val:,}",
                            "검사 수량": f"{total_qty:,}", "양품수량": f"{st.session_state.good_qty:,}", "양품 수량(전/배 포함)": f"{good_include_front_rear:,}", "불량수량": f"{bad_qty:,}",
                            "양품률": f"{rate_good:.1f}%", "양품율(전/배 포함)": f"{rate_good_inc:.1f}%",
                            "완전불량률": f"{comp_rate_num:.1f}%", "전면불량률": f"{front_rate_num:.1f}%", "배면불량률": f"{rear_rate_num:.1f}%",
                            "완전불량": f"{st.session_state.comp_def:,}", "전면불량": f"{st.session_state.front_def:,}", "배면불량": f"{st.session_state.rear_def:,}", "옵셋불량": f"{st.session_state.offset_def:,}", "수량부족": f"{st.session_state.shortage_qty:,}", "기타": f"{st.session_state.etc_def:,}",
                            "OQC": "" if st.session_state.oqc_status == "선택안함" else st.session_state.oqc_status, 
                            "비고": st.session_state.remarks, "도장라인": fmt_paint_line, "도장일": fmt_paint_date, 
                            "도장순서": st.session_state.painting_order, "입고일": fmt_in_date, "LOT NO.": st.session_state.lot_input_field, 
                            "CLIP": "" if st.session_state.clip_val == "선택안함" else st.session_state.clip_val, 
                            "BASE": "" if st.session_state.base_val == "선택안함" else st.session_state.base_val, 
                            "COVER": "" if st.session_state.cover_val == "선택안함" else st.session_state.cover_val, 
                            "조립기": fmt_assembler, "월": f"{st.session_state.work_date.month}월", 
                            "작업자": fmt_worker
                        }])
                        
                        if save_data_append(new_data):
                            st.success("저장 완료!")
                            for k, v in default_state.items():
                                st.session_state[k] = v
                            time.sleep(1)
                            st.session_state.step = 6
                            st.rerun()

    elif step == 6:
        st.markdown("---")
        df_history = load_data().copy()
        if not df_history.empty:
            recent_10 = df_history.iloc[::-1].head(10).copy()
            st.markdown("**최근 저장 데이터 List**")
            st.dataframe(recent_10, use_container_width=True, hide_index=True)
        else:
            st.caption("저장된 데이터가 없습니다.")
