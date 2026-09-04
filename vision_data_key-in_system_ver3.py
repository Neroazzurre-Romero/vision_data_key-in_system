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

# [설정] 작업자 명단
worker_list = ["박경섭", "무고사", "재르소", "김동헌"] 
model_list = ["D65S(KRIOS)", "MEM", "Centaur", "Sphinx-E", "Banff", "AV-J", "Seattle", "Juliet-O"]

st.set_page_config(
    page_title="VISION DATA KEY-IN SYSTEM ----- (by. Romero)", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# 영구 세션 상태 초기화 (페이지 이동 시 데이터 증발 방지)
# ==========================================
if "unlocked" not in st.session_state: st.session_state.unlocked = False
if "current_page" not in st.session_state: st.session_state.current_page = "input"
if "step" not in st.session_state: st.session_state.step = 1

default_state = {
    "work_date": datetime.now().date(), "shift_type": "주간", "worker_name": "선택안함", 
    "model_name": "D65S(KRIOS)", "lot_input_field": "", "in_date_field": datetime.now().date(),
    "plating_type": "A", "start_date": datetime.now().date(), "start_time": datetime.now().time(),
    "end_date": datetime.now().date(), "end_time": datetime.now().time(), "unit": "1호기",
    "category": "1차 검사", "idle_time": 0, "painting_date": datetime.now().date(),
    "painting_order": 1, "painting_line": "선택안함", "clip_val": "선택안함",
    "base_val": "선택안함", "cover_val": "선택안함", "assembler_val": "선택안함",
    "good_qty": 0, "comp_def": 0, "front_def": 0, "rear_def": 0, "offset_def": 0,
    "shortage_qty": 0, "etc_def": 0, "oqc_status": "선택안함", "remarks": ""
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
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
    </style>
    """
    st.markdown(hide_sidebar_style, unsafe_allow_html=True)
    
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    
    # 💡 로고 영역의 열 비율을 [1, 0.75, 1]로 조정하여 크기를 기존 대비 정확히 절반으로 축소
    c1, c2, c3 = st.columns([1, 0.75, 1])
    with c2:
        if os.path.exists("logo.png"):
            st.image("logo.png", use_container_width=True)
        else:
            st.error("⚠️ 'logo.png' 파일이 파이썬 실행 폴더에 없습니다. 이미지를 같은 폴더에 추가해주세요.")
            st.markdown("<h1 style='text-align: center; color: #1e293b; font-size: 60px; font-weight: 900;'>COMPANY LOGO</h1>", unsafe_allow_html=True)
        
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
                for(let b of btns) {
                    if(b.innerText.includes('UNLOCK_SYSTEM_BTN_HIDDEN')) {
                        b.click();
                        break;
                    }
                }
            };

            const btns = window.parent.document.querySelectorAll('button');
            for(let b of btns) {
                if(b.innerText.includes('UNLOCK_SYSTEM_BTN_HIDDEN')) {
                    b.style.display = 'none';
                }
            }

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
                
                if (currentX > maxDrag * 0.4) {
                    text.style.color = '#ffffff';
                } else {
                    text.style.color = '#94a3b8';
                }

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
                    setTimeout(() => {
                        thumb.style.transition = 'none';
                        fill.style.transition = 'none';
                    }, 300);
                }
            }

            thumb.addEventListener('mousedown', startDrag);
            document.addEventListener('mousemove', drag);
            document.addEventListener('mouseup', endDrag);

            thumb.addEventListener('touchstart', startDrag, {passive: false});
            document.addEventListener('touchmove', drag, {passive: false});
            document.addEventListener('touchend', endDrag);
        </script>
        """
        components.html(slider_html, height=90)
            
    st.markdown(
        "<div style='position: fixed; bottom: 10%; left: 0; width: 100%; text-align: center; font-size: 10pt; color: #94a3b8; font-weight: bold;'>vision data key-in system --- Romero.K</div>", 
        unsafe_allow_html=True
    )
    st.stop()

# ----------------------------------------------------
# 마법 코드 1: UI 디자인 커스텀 및 사이드바 스타일링
# ----------------------------------------------------
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
body { overscroll-behavior-y: none !important; }
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

[data-testid="stSidebar"] {
    background: linear-gradient(135deg, #0f172a 0%, #020617 100%) !important;
}
[data-testid="stSidebar"] * {
    color: #f8fafc !important;
}

[data-testid="stSidebar"] .stButton > button {
    height: 65px !important; 
    justify-content: flex-start !important; 
    padding-left: 15px !important; 
    margin-bottom: 6px !important;
    border-radius: 8px !important;
    background-color: transparent !important;
    color: #f8fafc !important;
    border: 1px solid #334155 !important;
}
[data-testid="stSidebar"] .stButton > button p {
    font-weight: 800 !important;
    font-size: 16px !important;
    text-indent: 10px !important; 
    text-align: left !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background-color: #3b82f6 !important; 
    color: white !important;
    border: 1px solid #2563eb !important;
}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

components.html(
    """
    <script>
    if (window.parent && !window.parent.appPluginLoaded) {
        window.parent.appPluginLoaded = true;
        
        const formatNavButtons = () => {
            if (!window.parent.document) return;
            const buttons = window.parent.document.querySelectorAll('button');
            buttons.forEach(btn => {
                const text = btn.innerText || "";
                if (text.includes('⬅️ 이전')) {
                    btn.style.backgroundColor = '#FFC000';
                    btn.style.color = '#000000';
                    btn.style.border = 'none';
                }
                if (text.includes('다음 ➡️')) {
                    btn.style.backgroundColor = '#00B050';
                    btn.style.color = '#FFFFFF';
                    btn.style.border = 'none';
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

        const observer = new MutationObserver(() => { 
            disableKeyboard(); 
            formatNavButtons();
        });
        
        if (window.parent.document.body) {
            observer.observe(window.parent.document.body, { childList: true, subtree: true });
        }
        
        disableKeyboard();
        formatNavButtons();
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

# ==========================================
# 메인 프로세스 화면 구성
# ==========================================
if st.session_state.current_page == "analysis":
    st.markdown("## 종합 생산 데이터 분석")
    if st.button("뒤로 가기 (데이터 입력 화면으로)"):
        st.session_state.current_page = "input"
        st.rerun()
    df = load_data().copy()
    if df.empty:
        st.warning("데이터가 없습니다.")
    else:
        st.dataframe(df.head(50))

elif st.session_state.current_page == "input":
    st.markdown("""
        <div style='background: linear-gradient(135deg, #0f172a 0%, #020617 100%); padding: 10px 20px; border-radius: 8px; margin-bottom: 15px; box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.5); border: 1px solid #1e293b;'>
            <h2 style='color: #f8fafc; margin: 0; font-weight: 600;'>VISION DATA KEY-IN SYSTEM</h2>
        </div>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("### 📋 목차")
        steps_titles = [
            "생산 등록", "작업 정보", "Coating Data", "Assemble Data", 
            "VISION Data & ETC", "데이터 저장", "Report & History"
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
            if st.session_state.step < 7:
                if st.button("다음 ➡️", use_container_width=True):
                    st.session_state.step += 1
                    st.rerun()

    step = st.session_state.step

    if step == 1:
        st.markdown("### 1. 생산 등록")
        
        c1, c2 = st.columns(2)
        with c1: st.session_state.work_date = st.date_input("**근무일자**", value=st.session_state.work_date)
        with c2: st.session_state.model_name = st.selectbox("**모델명**", model_list, index=model_list.index(st.session_state.model_name) if st.session_state.model_name in model_list else 0)
        
        st.markdown("**교대**")
        render_grid_buttons(["주간", "야간"], "shift_type", 2)
        
        st.session_state.worker_name = st.selectbox("**작업자**", ["선택안함"] + worker_list, index=(["선택안함"] + worker_list).index(st.session_state.worker_name) if st.session_state.worker_name in (["선택안함"] + worker_list) else 0)
        
        def parse_scanned_data():
            raw_val = st.session_state._lot_input_temp
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

        st.info("💡 스캐너 키보드 앱으로 바코드를 스캔하면 도금구분, 입고일, LOT 번호가 자동으로 분류됩니다.")
        st.text_input("**LOT 직접 입력 및 스캔**", value=st.session_state.lot_input_field, key="_lot_input_temp", on_change=parse_scanned_data, placeholder="입력창 터치 후 스캔")
        
        st.session_state.in_date_field = st.date_input("**입고일 (자동세팅됨)**", value=st.session_state.in_date_field)
        
        st.markdown("**도금 구분 (자동세팅됨)**")
        render_grid_buttons(["A", "B"], "plating_type", 2)

    elif step == 2:
        st.markdown("### 2. 작업 정보")
        
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
        st.markdown("### 3. Coating Data")
        c1, c2 = st.columns(2)
        with c1: st.session_state.painting_date = st.date_input("**도장일**", value=st.session_state.painting_date)
        with c2: st.session_state.painting_order = st.number_input("**도장순서**", min_value=1, value=st.session_state.painting_order)
        
        st.markdown("**도장라인**")
        render_grid_buttons(["A Line", "B Line", "C Line"], "painting_line", 3)

    elif step == 4:
        st.markdown("### 4. Assemble Data")
        num_options = ["선택안함"] + [str(i) for i in range(1, 11)]
        
        c1, c2, c3 = st.columns(3)
        with c1: st.session_state.clip_val = st.selectbox("**CLIP**", num_options, index=num_options.index(st.session_state.clip_val) if st.session_state.clip_val in num_options else 0)
        with c2: st.session_state.base_val = st.selectbox("**BASE**", num_options, index=num_options.index(st.session_state.base_val) if st.session_state.base_val in num_options else 0)
        with c3: st.session_state.cover_val = st.selectbox("**COVER**", num_options, index=num_options.index(st.session_state.cover_val) if st.session_state.cover_val in num_options else 0)

        st.markdown("**조립기**")
        render_grid_buttons(["1호기", "2호기", "3호기", "4호기"], "assembler_val", 2)

    elif step == 5:
        st.markdown("### 5. VISION Data & ETC")
        
        st.session_state.good_qty = st.number_input("**양품수량**", min_value=0, value=st.session_state.good_qty)
        
        st.markdown("##### 🚨 불량 세부")
        c1, c2, c3 = st.columns(3)
        with c1: st.session_state.comp_def = st.number_input("**완전불량**", min_value=0, value=st.session_state.comp_def)
        with c2: st.session_state.front_def = st.number_input("**전면불량**", min_value=0, value=st.session_state.front_def)
        with c3: st.session_state.rear_def = st.number_input("**배면불량**", min_value=0, value=st.session_state.rear_def)
        
        c4, c5, c6 = st.columns(3)
        with c4: st.session_state.offset_def = st.number_input("**옵셋불량**", min_value=0, value=st.session_state.offset_def)
        with c5: st.session_state.shortage_qty = st.number_input("**수량부족**", min_value=0, value=st.session_state.shortage_qty)
        with c6: st.session_state.etc_def = st.number_input("**기타**", min_value=0, value=st.session_state.etc_def)
        
        bad_qty = st.session_state.comp_def + st.session_state.front_def + st.session_state.rear_def + st.session_state.offset_def + st.session_state.etc_def
        total_qty = max(0, st.session_state.good_qty + bad_qty - st.session_state.shortage_qty)
        
        st.markdown("##### 합계")
        h1, h2 = st.columns(2)
        with h1: st.text_input("**검사 수량 (자동)**", value=f"{total_qty:,}", disabled=True)
        with h2: st.text_input("**불량수량 (자동)**", value=f"{bad_qty:,}", disabled=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.session_state.oqc_status = st.selectbox("**OQC**", ["선택안함", "육안", "OQC"], index=["선택안함", "육안", "OQC"].index(st.session_state.oqc_status))
        st.session_state.remarks = st.text_area("**비고**", value=st.session_state.remarks, height=68)

    elif step == 6:
        st.markdown("### 6. 데이터 저장")
        
        bad_qty = st.session_state.comp_def + st.session_state.front_def + st.session_state.rear_def + st.session_state.offset_def + st.session_state.etc_def
        total_qty = max(0, st.session_state.good_qty + bad_qty - st.session_state.shortage_qty)
        
        st.info("입력된 데이터를 구글 스프레드시트에 전송합니다. 마지막으로 확인 후 아래 저장 버튼을 눌러주세요.")
        st.write(f"- **LOT NO:** {st.session_state.lot_input_field}")
        st.write(f"- **검사 총 수량:** {total_qty:,} 개")
        
        if st.button("💾 데이터 최종 저장 (구글 시트 전송)", type="primary", use_container_width=True):
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
                        "작업자": "" if st.session_state.worker_name == "선택안함" else st.session_state.worker_name
                    }])
                    if save_data_append(new_data):
                        st.success("저장 완료!")
                        st.session_state.lot_input_field = ""
                        time.sleep(1)
                        st.session_state.step = 7
                        st.rerun()

    elif step == 7:
        st.markdown("### 7. Report & History")
        
        if st.button("📊 종합 분석 데이터 페이지로 이동", use_container_width=True, type="primary"):
            st.session_state.current_page = "analysis"
            st.rerun()

        st.markdown("---")
        bad_qty = st.session_state.comp_def + st.session_state.front_def + st.session_state.rear_def + st.session_state.offset_def + st.session_state.etc_def
        total_qty = max(0, st.session_state.good_qty + bad_qty - st.session_state.shortage_qty)
        rate_good = round((st.session_state.good_qty / total_qty) * 100, 1) if total_qty > 0 else 0.0

        m_col1, m_col2 = st.columns(2)
        with m_col1: st.metric(label="최근 검사수량 총합", value=f"{total_qty:,} EA")
        with m_col2: st.metric(label="최근 양품율", value=f"{rate_good:.1f}%")

        df_history = load_data().copy()
        if not df_history.empty:
            recent_10 = df_history.iloc[::-1].head(10).copy()
            st.markdown("**최근 저장 데이터 List**")
            st.dataframe(recent_10, use_container_width=True, hide_index=True)
        else:
            st.caption("저장된 데이터가 없습니다.")
