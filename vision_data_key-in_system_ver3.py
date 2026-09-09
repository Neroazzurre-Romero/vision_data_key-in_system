import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
import os
from datetime import datetime, time as dt_time
import time
from io import BytesIO
from openpyxl.styles import Font
import openpyxl
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials
import base64

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

st.set_page_config(page_title="VISION DATA KEY-IN SYSTEM", layout="wide", initial_sidebar_state="expanded")

def get_image_base64(base_name):
    search_dirs = [os.getcwd(), os.path.dirname(os.path.abspath(__file__))]
    for directory in search_dirs:
        if not os.path.exists(directory): continue
        for file in os.listdir(directory):
            if file.lower().startswith(base_name.lower()) and file.lower().endswith(('.png', '.jpg', '.jpeg')):
                filepath = os.path.join(directory, file)
                try:
                    with open(filepath, "rb") as img_file:
                        ext = file.split('.')[-1].lower()
                        mime_type = "image/jpeg" if ext in ['jpg', 'jpeg'] else "image/png"
                        encoded = base64.b64encode(img_file.read()).decode('utf-8')
                        return f"data:{mime_type};base64,{encoded}"
                except Exception:
                    pass
    return None

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
    "rear_warned": False, "offset_warned": False,
    "numpad_buffer": "",
    "timepad_buffer": ""
}

for key, value in default_state.items():
    if key not in st.session_state:
        st.session_state[key] = value

if not st.session_state.unlocked:
    hide_sidebar_style = """
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        [data-testid="stToolbar"] { display: none !important; }
        footer { display: none !important; }
        [data-testid="stAppViewContainer"] { background-color: #ffffff !important; }
    </style>
    """
    st.markdown(hide_sidebar_style, unsafe_allow_html=True)
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        logo_l_data = get_image_base64("logo")
        if logo_l_data:
            st.markdown(f"<div style='text-align: center;'><img src='{logo_l_data}' style='max-width: 100%; max-height: 180px; object-fit: contain; margin-bottom: 20px;'></div>", unsafe_allow_html=True)
        else:
            st.markdown("<h1 style='text-align: center; color: #1e293b; font-size: 45px; font-weight: 900; letter-spacing: 2px;'>VISION DATA KEY-IN SYSTEM</h1><br><br>", unsafe_allow_html=True)
        
        if st.button("UNLOCK_SYSTEM_BTN_HIDDEN"):
            st.session_state.unlocked = True
            st.rerun()

        slider_html = """
        <div id="slider-container" style="background: #ffffff; border: 2px solid #e2e8f0; border-radius: 40px; position: relative; width: 100%; max-width: 400px; height: 68px; margin: 0 auto; overflow: hidden; display: flex; align-items: center; box-shadow: inset 0 2px 5px rgba(0,0,0,0.05);">
            <div id="slider-fill" style="position: absolute; left: 0; top: 0; height: 100%; width: 0; background-color: #1e293b; border-radius: 40px 0 0 40px;"></div>
            <div id="slider-text" style="position: absolute; width: 100%; text-align: center; color: #94a3b8; font-size: 20px; font-weight: bold; font-family: sans-serif; pointer-events: none; z-index: 2; transition: color 0.3s;">Slide to Unlock</div>
            <div id="slider-thumb" style="position: absolute; left: 4px; width: 56px; height: 56px; background: #ffffff; border-radius: 50%; box-shadow: 0 2px 6px rgba(0,0,0,0.2); cursor: pointer; z-index: 3; display: flex; align-items: center; justify-content: center; color: #1e293b; font-size: 24px;">▶</div>
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
            
    st.markdown("<div style='position: fixed; bottom: 10%; left: 0; width: 100%; text-align: center; font-size: 10pt; color: #94a3b8; font-weight: bold;'>Create by --- Romero.K</div>", unsafe_allow_html=True)
    st.stop()

# ----------------------------------------------------
# 💡 단색(Flat) 테마 & 입력창/버튼 동적 색상 변경 CSS
# ----------------------------------------------------
hide_streamlit_style = """
<style>
footer { display: none !important; } 
[data-testid="collapsedControl"] { display: flex !important; visibility: visible !important; opacity: 1 !important; z-index: 99999 !important; }
body { overscroll-behavior-y: none !important; } 
::-webkit-scrollbar { display: none; }
.block-container { padding-top: 3rem !important; padding-bottom: 2rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 95% !important; }

/* 배경 및 카드 레이아웃 */
[data-testid="stAppViewContainer"] { background-color: #f1f5f9 !important; }
[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #ffffff !important;
    border-radius: 12px !important;
    border: 1px solid #cbd5e1 !important;
    box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04) !important;
    padding: 1.5rem !important;
    margin-bottom: 0.5rem !important;
}

div[data-testid="stMarkdownContainer"] p strong { font-size: 1.1rem !important; font-weight: 800 !important; color: #1e293b !important; }

/* 공통 높이(2.6rem - 2/3 사이즈) */
div[data-testid="stButton"] button { 
    height: 2.6rem !important; 
    min-height: 2.6rem !important; 
    max-height: 2.6rem !important;
    font-size: 1.1rem !important; 
    font-weight: bold !important; 
    border-radius: 8px !important;
    width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
    box-sizing: border-box !important;
    transition: all 0.2s ease;
}

/* 💡 입력창 기본 테마 (배경 #E7E6E6, 텍스트 검정) */
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div {
    height: 2.6rem !important;
    min-height: 2.6rem !important;
    max-height: 2.6rem !important;
    border-radius: 8px !important;
    background-color: #E7E6E6 !important;
    border: 1px solid #cbd5e1 !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
    box-sizing: border-box !important;
    transition: all 0.2s ease;
}

/* 💡 텍스트 색상 (기본 검정) */
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div > div,
div[data-testid="stDateInput"] input,
div[data-testid="stTextInput"] input {
    height: 2.6rem !important;
    min-height: 2.6rem !important;
    max-height: 2.6rem !important;
    line-height: 2.6rem !important;
    font-size: 1.1rem !important;
    font-weight: bold !important;
    text-align: center !important;
    color: #000000 !important; 
    padding: 0 10px !important;
    margin: 0 !important;
    background: transparent !important;
    border: none !important;
    box-sizing: border-box !important;
    transition: color 0.2s ease;
}

div[data-testid="stSelectbox"] div[data-baseweb="select"] > div > div:last-child {
    display: flex !important;
    align-items: center !important;
    height: 2.6rem !important;
}

/* 💡 비고란 (Textarea) - 최종 저장 버튼과 동일한 150px 높이 강제 설정 */
div[data-baseweb="textarea"] textarea { 
    font-size: 1.1rem !important; 
    height: 150px !important;
    min-height: 150px !important; 
    max-height: 150px !important; 
    background-color: #E7E6E6 !important; 
    color: #000000 !important;
    border: 1px solid #cbd5e1 !important; 
    border-radius: 8px !important; 
    padding: 15px !important;
    transition: all 0.2s ease;
}

/* 💡 입력창 포커스 시 적용버튼 색상(#1e293b) + 글자 흰색 전환 */
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div:focus-within,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div:focus-within {
    background-color: #1e293b !important;
    border-color: #1e293b !important;
}

div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within > div,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div:focus-within input,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div:focus-within input {
    color: #ffffff !important;
}

div[data-baseweb="textarea"]:focus-within textarea {
    background-color: #1e293b !important;
    color: #ffffff !important;
    border-color: #1e293b !important;
}

div[data-baseweb="select"] input, div[data-baseweb="datepicker"] input {
    caret-color: transparent !important;
    cursor: pointer !important;
}

/* 스캐너 Placeholder 강제 스타일 */
input[placeholder*="스캐너 앱 실행"] { color: #000000 !important; font-weight: 900 !important; }
input[placeholder*="스캐너 앱 실행"]::placeholder { color: #4b5563 !important; font-weight: bold !important; opacity: 0.8 !important; }

/* 💡 일반 버튼(Secondary) 단색 스타일 (배경 #E7E6E6, 텍스트 검정) */
div[data-testid="stButton"] button[kind="secondary"] {
    background-color: #E7E6E6 !important;
    color: #000000 !important;
    border: 1px solid #cbd5e1 !important;
    box-shadow: none !important;
}
/* 포커스/호버 시 적용버튼 색상으로 변경 */
div[data-testid="stButton"] button[kind="secondary"]:hover,
div[data-testid="stButton"] button[kind="secondary"]:focus,
div[data-testid="stButton"] button[kind="secondary"]:active {
    background-color: #1e293b !important;
    color: #ffffff !important;
    border-color: #1e293b !important;
}

/* 💡 Primary 버튼 단색 스타일 (차콜 네이비) */
div[data-testid="stButton"] button[kind="primary"] {
    background-color: #1e293b !important;
    color: #ffffff !important;
    border: 1px solid #0f172a !important;
    box-shadow: none !important;
}
div[data-testid="stButton"] button[kind="primary"]:hover {
    background-color: #0f172a !important;
}

/* 사이드바 메뉴 단색 디자인 */
[data-testid="stSidebar"] { background-color: #0f172a !important; }
[data-testid="stSidebar"] * { color: #f8fafc !important; }
[data-testid="stSidebar"] .stButton > button { 
    height: 80px !important; 
    max-height: 80px !important;
    justify-content: flex-start !important; 
    padding-left: 15px !important; 
    margin-bottom: 10px !important; 
    border-radius: 6px !important; 
    background-color: transparent !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button p { font-weight: 800 !important; font-size: 16px !important; text-indent: 10px !important; text-align: left !important; }

/* 선택된 사이드바 버튼 */
[data-testid="stSidebar"] .stButton > button[kind="primary"] { 
    background: linear-gradient(90deg, #000000 0%, #0033A0 100%) !important; 
    color: #FFFFFF !important; 
    border: none !important; 
    border-left: 4px solid #FFC000 !important;
    box-shadow: none !important;
}
/* 미선택 사이드바 버튼 */
[data-testid="stSidebar"] .stButton > button[kind="secondary"] { 
    background-color: transparent !important; 
    color: #8B9CB6 !important; 
    border: 1px solid transparent !important; 
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
    background-color: transparent !important;
    color: #ffffff !important;
    border: 1px solid transparent !important;
    box-shadow: none !important;
}
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
                
                if (text.includes('⬅️ 이전')) { 
                    btn.style.backgroundColor = '#E7E6E6'; 
                    btn.style.background = 'none';
                    btn.style.color = '#000000'; 
                    btn.style.border = '1px solid #cbd5e1'; 
                    btn.style.setProperty('height', '65px', 'important'); 
                }
                if (text.includes('다음 ➡️')) { 
                    btn.style.backgroundColor = '#16a34a'; 
                    btn.style.background = 'none';
                    btn.style.color = '#ffffff'; 
                    btn.style.border = '1px solid #15803d'; 
                    btn.style.setProperty('height', '65px', 'important'); 
                }
                if (text.includes('데이터 최종 저장')) { 
                    btn.style.backgroundColor = '#0369a1';
                    btn.style.background = 'none';
                    btn.style.border = '1px solid #0c4a6e';
                    btn.style.color = '#ffffff';
                    btn.style.setProperty('height', '150px', 'important');
                    btn.style.setProperty('max-height', '150px', 'important');
                    btn.style.setProperty('margin-top', '0px', 'important'); 
                    btn.style.setProperty('font-size', '20px', 'important');
                    btn.style.setProperty('white-space', 'pre-wrap', 'important');
                }
                // 💡 Data Analysis 버튼 높이를 메인 헤더박스와 동일하게(6rem) 2배 확장
                if (text.trim() === 'Data Analysis') { 
                    btn.style.backgroundColor = '#1e293b';
                    btn.style.background = 'none';
                    btn.style.color = '#ffffff';
                    btn.style.border = '1px solid #0f172a';
                    btn.style.boxShadow = 'none';
                    btn.style.setProperty('height', '6rem', 'important');
                    btn.style.setProperty('min-height', '6rem', 'important');
                    btn.style.setProperty('max-height', '6rem', 'important');
                    btn.style.setProperty('font-size', '1.6rem', 'important');
                    btn.style.setProperty('margin-top', '0px', 'important');
                }
            });
        };
        
        // 💡 스캐너 텍스트칸 동적 포커스 및 노란색 강조
        const styleScanner = () => {
            if (!window.parent.document) return;
            window.parent.document.querySelectorAll('input').forEach(el => {
                if (el.getAttribute('placeholder') && el.getAttribute('placeholder').includes('스캐너 앱 실행')) {
                    el.style.setProperty('font-size', '1.2rem', 'important');
                    el.style.setProperty('font-weight', '900', 'important');
                    
                    let parentDiv = el.parentElement;
                    let grandParent = el.closest('div[data-baseweb="input"]');
                    if (parentDiv) parentDiv.style.setProperty('border', 'none', 'important');
                    if (grandParent) grandParent.style.setProperty('border', '2px solid #eab308', 'important');
                    
                    // 이벤트 리스너를 통해 포커스 상태 감지하여 색상 반전
                    if (!el.getAttribute('data-scanner-listener')) {
                        el.setAttribute('data-scanner-listener', 'true');
                        el.addEventListener('focus', () => { el.setAttribute('data-focused', 'true'); });
                        el.addEventListener('blur', () => { el.removeAttribute('data-focused'); });
                    }
                    
                    if (el.getAttribute('data-focused')) {
                        el.style.setProperty('color', '#ffffff', 'important');
                        if (parentDiv) parentDiv.style.setProperty('background-color', '#1e293b', 'important');
                        if (grandParent) grandParent.style.setProperty('background-color', '#1e293b', 'important');
                    } else {
                        el.style.setProperty('color', '#000000', 'important');
                        if (parentDiv) parentDiv.style.setProperty('background-color', '#fef08a', 'important');
                        if (grandParent) grandParent.style.setProperty('background-color', '#fef08a', 'important');
                    }
                }
            });
        };
        
        const disableKeyboard = () => {
            if (!window.parent.document) return;
            const inputs = window.parent.document.querySelectorAll('input');
            inputs.forEach(el => {
                const placeholder = el.getAttribute('placeholder') || '';
                const ariaLabel = el.getAttribute('aria-label') || '';
                const isDropdown = el.closest('div[data-baseweb="select"]') !== null;
                const isDatepicker = el.closest('div[data-baseweb="datepicker"]') !== null;
                
                if (placeholder.includes('YYYY') || placeholder.includes('MM') || placeholder.includes('DD') || 
                    ariaLabel.toLowerCase().includes('date') || ariaLabel.toLowerCase().includes('select') || 
                    isDropdown || isDatepicker) {
                    
                    el.setAttribute('inputmode', 'none');
                    el.setAttribute('readonly', 'readonly');
                    el.addEventListener('focus', function(e) {
                        e.target.blur();
                    });
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

# 💡 교대 주간/야간, 도금구분이 폭을 동일하게 채울 수 있도록 use_width=True로 세팅
def render_grid_buttons(options, state_key, columns, use_width=True):
    rows = [options[i:i+columns] for i in range(0, len(options), columns)]
    for row_opts in rows:
        cols = st.columns(columns)
        for i, opt in enumerate(row_opts):
            with cols[i]:
                if opt.strip() == "": st.write("") 
                else:
                    btn_type = "primary" if st.session_state[state_key] == opt else "secondary"
                    if st.button(opt, key=f"btn_{state_key}_{opt}", type=btn_type, use_container_width=use_width):
                        st.session_state[state_key] = opt
                        st.rerun()

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

@st.dialog("📷 바코드/QR 스캐너")
def scanner_dialog():
    st.markdown("<div style='text-align:center; font-size:1.2rem; font-weight:bold; color:#155724; padding:15px; background:#fef08a; border-radius:10px; margin-bottom:15px; border:2px solid #eab308;'>아래 입력창을 터치하여 스캐너 앱을 띄운 후 스캔하세요.</div>", unsafe_allow_html=True)
    
    raw_scan = st.text_input("바코드 데이터", key="dialog_scan_input", label_visibility="collapsed", placeholder="여기를 터치하여 스캔하세요")
    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("적용 (Enter)", type="primary", use_container_width=True):
        if raw_scan:
            st.session_state.scanned_raw_data = raw_scan
            parts = [p for p in raw_scan.split('$') if p]
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
                st.session_state.lot_input_field = parts[-1] if '$' in raw_scan else raw_scan
        st.rerun()

def pad_callback(digit):
    c_val = st.session_state.numpad_buffer
    if digit == "C": 
        st.session_state.numpad_buffer = ""
    elif digit == "⬅": 
        st.session_state.numpad_buffer = c_val[:-1]
    else:
        if len(c_val) < 8: 
            st.session_state.numpad_buffer = c_val + digit

@st.dialog("🔢 수량 입력 패드")
def numpad_dialog(field_key, display_name):
    c_val = st.session_state.numpad_buffer
    st.markdown(f"<div style='text-align:center; font-size:1.8rem; font-weight:bold; color:#1e293b; padding:15px; background:#f8fafc; border-radius:10px; margin-bottom:15px; border:1px solid #cbd5e1;'>{display_name}<br><span style='color:#1e293b; font-size:2.5rem;'>{int(c_val) if c_val else 0:,}</span></div>", unsafe_allow_html=True)
    
    pad_rows = [
        ["7", "8", "9"],
        ["4", "5", "6"],
        ["1", "2", "3"],
        ["C", "0", "⬅"]
    ]
    
    for r in pad_rows:
        cols = st.columns(3)
        for i, val in enumerate(r):
            with cols[i]:
                st.button(val, key=f"pad_{field_key}_{val}", use_container_width=True, on_click=pad_callback, args=(val,))
                    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("적용 (Enter)", type="primary", use_container_width=True):
        st.session_state[field_key] = int(st.session_state.numpad_buffer) if st.session_state.numpad_buffer else 0
        st.session_state.numpad_buffer = "" 
        st.rerun()

def timepad_callback(digit):
    c_val = st.session_state.timepad_buffer
    if digit == "C": 
        st.session_state.timepad_buffer = ""
    elif digit == "⬅": 
        st.session_state.timepad_buffer = c_val[:-1]
    else:
        if len(c_val) < 4: 
            st.session_state.timepad_buffer = c_val + digit

@st.dialog("⏰ 시간 입력 패드 (HH:MM)")
def timepad_dialog(field_key, display_name):
    c_val = st.session_state.timepad_buffer
    display_str = c_val.ljust(4, "_")
    display_str = f"{display_str[:2]}:{display_str[2:]}"
    
    st.markdown(f"<div style='text-align:center; font-size:1.5rem; font-weight:bold; color:#1e293b; padding:15px; background:#f8fafc; border-radius:10px; margin-bottom:15px; border:1px solid #cbd5e1;'>{display_name}<br><span style='color:#1e293b; font-size:2.5rem; letter-spacing: 2px;'>{display_str}</span></div>", unsafe_allow_html=True)
    
    pad_rows = [
        ["7", "8", "9"],
        ["4", "5", "6"],
        ["1", "2", "3"],
        ["C", "0", "⬅"]
    ]
    
    for r in pad_rows:
        cols = st.columns(3)
        for i, val in enumerate(r):
            with cols[i]:
                st.button(val, key=f"tpad_{field_key}_{val}", use_container_width=True, on_click=timepad_callback, args=(val,))
                    
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("적용 (Enter)", type="primary", use_container_width=True):
        if len(st.session_state.timepad_buffer) == 4:
            try:
                h = int(st.session_state.timepad_buffer[:2])
                m = int(st.session_state.timepad_buffer[2:])
                if 0 <= h <= 23 and 0 <= m <= 59:
                    st.session_state[field_key] = dt_time(h, m)
                    st.session_state.timepad_buffer = "" 
                    st.rerun()
                else:
                    st.error("유효한 시간(00~23)과 분(00~59)을 입력하세요.")
            except ValueError:
                pass
        else:
            st.error("4자리 숫자를 모두 입력하세요 (예: 0830)")

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
    logo_s_data = get_image_base64("at")
    img_html = f"<img src='{logo_s_data}' style='height: 40px; margin-right: 15px; vertical-align: middle;'>" if logo_s_data else ""
    st.markdown(f"<h2 style='display: flex; align-items: center; color: #1e293b;'>{img_html} 종합 생산 데이터 분석 📊</h2>", unsafe_allow_html=True)
    
    col1, col2 = st.columns([0.8, 0.2])
    with col2:
        if st.button("돌아가기 (데이터 입력)", type="primary", use_container_width=True):
            st.session_state.current_page = "input"
            st.rerun()
            
    df = load_data().copy()
    if df.empty: 
        st.warning("저장된 데이터가 없습니다.")
    else:
        numeric_cols = ["검사 수량", "양품수량", "불량수량", "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타"]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)

        with st.container(border=True):
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 일자별 양/불량 현황</h4>", unsafe_allow_html=True)
                df_date = df.groupby('날짜')[['양품수량', '불량수량']].sum().reset_index()
                fig1 = px.bar(df_date, x='날짜', y=['양품수량', '불량수량'], barmode='group', 
                              color_discrete_sequence=['#10B981', '#EF4444'])
                fig1.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#1e293b'))
                st.plotly_chart(fig1, use_container_width=True)
                
            with g_col2:
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 주요 불량 유형 비율</h4>", unsafe_allow_html=True)
                defect_sums = df[['완전불량', '전면불량', '배면불량', '옵셋불량', '기타']].sum()
                fig2 = px.pie(names=defect_sums.index, values=defect_sums.values, hole=0.5, 
                              color_discrete_sequence=['#EF4444', '#F59E0B', '#1e293b', '#8B5CF6', '#6B7280'])
                fig2.update_layout(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', font=dict(color='#1e293b'))
                st.plotly_chart(fig2, use_container_width=True)
            
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 전체 데이터 내역</h4>", unsafe_allow_html=True)
            st.dataframe(df, use_container_width=True, hide_index=True)

elif st.session_state.current_page == "input":
    
    top_c1, top_c2 = st.columns([5, 1])
    with top_c1:
        logo_s_data = get_image_base64("at")
        img_html = f"<img src='{logo_s_data}' style='height: 60px; margin-right: 20px;'>" if logo_s_data else ""
        
        # 💡 헤더 박스 높이를 기존의 2배(6rem)로 확장 
        st.markdown(
            f"<div style='background: #ffffff; padding: 0 30px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #cbd5e1; height: 6rem; display: flex; align-items: center; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04); box-sizing: border-box;'>"
            f"{img_html}"
            f"<h3 style='color: #1e293b; margin: 0; font-weight: 900; font-size: 2.2rem; letter-spacing: 1px;'>VISION DATA KEY-IN SYSTEM</h3>"
            f"</div>", 
            unsafe_allow_html=True
        )
    with top_c2:
        if st.button("Data Analysis", use_container_width=True, type="primary"):
            st.session_state.current_page = "analysis"
            st.rerun()

    with st.sidebar:
        steps_titles = [
            "생산 등록", "작업 정보", "Assemble & Coating", 
            "VISION Data", "Report & History"
        ]
        for i, title in enumerate(steps_titles, 1):
            btn_type = "primary" if st.session_state.step == i else "secondary"
            if st.button(title, key=f"nav_btn_{i}", type=btn_type, use_container_width=True):
                st.session_state.step = i
                st.rerun()

        st.markdown("<hr style='border-color: #334155; margin-top: -5px; margin-bottom: 5px;'>", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            if st.session_state.step > 1:
                if st.button("⬅️ 이전", use_container_width=True):
                    st.session_state.step -= 1
                    st.rerun()
        with c2:
            if st.session_state.step < 5:
                if st.button("다음 ➡️", use_container_width=True):
                    st.session_state.step += 1
                    st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("<div style='text-align: center; color: #94a3b8; font-size: 14px; font-weight: bold;'>Create by --- Romero.K</div>", unsafe_allow_html=True)

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
        st.session_state.scanned_raw_data = "" 

    if step == 1:
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 기본 근무 정보</h4>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: 
                st.markdown("**근무일자**")
                st.session_state.work_date = st.date_input("근무일자", value=st.session_state.work_date, label_visibility="collapsed")
            with c2: 
                st.markdown("**모델명**")
                st.session_state.model_name = st.selectbox("모델명", model_list, index=model_list.index(st.session_state.model_name) if st.session_state.model_name in model_list else 0, label_visibility="collapsed")
            with c3:
                st.markdown("**교대**")
                render_grid_buttons(["주간", "야간"], "shift_type", 2, use_width=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("**작업자**")
            w_col1, w_col2, w_col3 = st.columns(3)
            with w_col1: st.session_state.worker_a = st.selectbox("A조", worker_a_list, index=worker_a_list.index(st.session_state.worker_a) if st.session_state.worker_a in worker_a_list else 0, label_visibility="collapsed")
            with w_col2: st.session_state.worker_b = st.selectbox("B조", worker_b_list, index=worker_b_list.index(st.session_state.worker_b) if st.session_state.worker_b in worker_b_list else 0, label_visibility="collapsed")
            with w_col3: st.session_state.worker_c = st.selectbox("C조", worker_c_list, index=worker_c_list.index(st.session_state.worker_c) if st.session_state.worker_c in worker_c_list else 0, label_visibility="collapsed")

        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 스캔 및 입고 정보</h4>", unsafe_allow_html=True)
            
            sc1, sc2, sc3 = st.columns(3)
            with sc1:
                scan_in, scan_btn = st.columns(2)
                with scan_in:
                    st.markdown("**스캔 데이터**")
                    st.text_input("스캔 데이터", key="scanned_raw_data", label_visibility="collapsed", placeholder="스캐너 앱 실행")
                with scan_btn:
                    st.markdown("**&nbsp;**")
                    st.button("적용", type="primary", use_container_width=True, on_click=parse_scanned_data)
            
            with sc2:
                lot_col, date_col = st.columns(2)
                with lot_col:
                    st.markdown("**LOT (적용됨)**")
                    st.text_input("LOT", value=st.session_state.lot_input_field, disabled=True, label_visibility="collapsed")
                with date_col:
                    st.markdown("**입고일 (적용됨)**")
                    st.date_input("입고일", value=st.session_state.in_date_field, disabled=True, label_visibility="collapsed")
            
            with sc3:
                st.markdown("**도금 구분**")
                plat_a, plat_b = st.columns(2)
                with plat_a:
                    btn_a = "primary" if st.session_state.plating_type == "A" else "secondary"
                    if st.button("A", type=btn_a, key="plat_a", use_container_width=True):
                        st.session_state.plating_type = "A"
                        st.rerun()
                with plat_b:
                    btn_b = "primary" if st.session_state.plating_type == "B" else "secondary"
                    if st.button("B", type=btn_b, key="plat_b", use_container_width=True):
                        st.session_state.plating_type = "B"
                        st.rerun()

    elif step == 2:
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 작업 시간</h4>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: 
                st.markdown("**시작일**")
                st.session_state.start_date = st.date_input("시작일", value=st.session_state.start_date, label_visibility="collapsed")
            with c2: 
                st.markdown("**시작시간**")
                time_str = st.session_state.start_time.strftime("%H:%M")
                if st.button(time_str, key="btn_start_time", use_container_width=True):
                    st.session_state.timepad_buffer = st.session_state.start_time.strftime("%H%M")
                    timepad_dialog("start_time", "시작시간")
            with c3: 
                st.markdown("**휴동시간 (분)**")
                if st.button(f"{st.session_state.idle_time:,}", key="btn_idle_time", use_container_width=True):
                    st.session_state.numpad_buffer = str(st.session_state.idle_time) if st.session_state.idle_time != 0 else ""
                    numpad_dialog("idle_time", "휴동시간 (분)")
            
            st.markdown("<br>", unsafe_allow_html=True)
            c4, c5, c6 = st.columns(3)
            with c4: 
                st.markdown("**종료일**")
                st.session_state.end_date = st.date_input("종료일", value=st.session_state.end_date, label_visibility="collapsed")
            with c5: 
                st.markdown("**종료시간**")
                time_str = st.session_state.end_time.strftime("%H:%M")
                if st.button(time_str, key="btn_end_time", use_container_width=True):
                    st.session_state.timepad_buffer = st.session_state.end_time.strftime("%H%M")
                    timepad_dialog("end_time", "종료시간")
            with c6: 
                st.markdown("**소요시간 (차감됨)**")
                start_dt = datetime.combine(st.session_state.start_date, st.session_state.start_time)
                end_dt = datetime.combine(st.session_state.end_date, st.session_state.end_time)
                raw_duration = int((end_dt - start_dt).total_seconds() / 60)
                duration_minutes = max(0, raw_duration - st.session_state.idle_time)
                st.text_input("소요시간", value=f"{duration_minutes:,} 분", disabled=True, label_visibility="collapsed")
        
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 설비 및 검사 설정</h4>", unsafe_allow_html=True)
            st.markdown("**호기**")
            render_grid_buttons(["1호기", "2호기", "3호기", "4호기", "5호기", "6호기"], "unit", 3)
            st.markdown("<br>**검사 구분**", unsafe_allow_html=True)
            render_grid_buttons(["1차 검사", "2차 검사", "3차 검사", "K 1차 검사", "Sample", "완불재검"], "category", 6)

    elif step == 3:
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 도장 공정</h4>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: 
                st.markdown("**도장일**")
                st.session_state.painting_date = st.date_input("도장일", value=st.session_state.painting_date, label_visibility="collapsed")
            with c2: 
                st.markdown("**도장라인**")
                render_grid_buttons(["A Line", "B Line", "C Line"], "painting_line", 3)
            with c3: 
                st.markdown("**도장순서**")
                if st.button(f"{st.session_state.painting_order:,}", key="btn_paint_order", use_container_width=True):
                    st.session_state.numpad_buffer = str(st.session_state.painting_order) if st.session_state.painting_order != 0 else ""
                    numpad_dialog("painting_order", "도장순서")
        
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ Assemble 부품 및 설비</h4>", unsafe_allow_html=True)
            num_options = ["1"] + [str(i) for i in range(2, 11)] + ["선택안함"]
            c4, c5, c6 = st.columns(3)
            with c4: 
                st.markdown("**CLIP**")
                st.session_state.clip_val = st.selectbox("CLIP", num_options, index=num_options.index(st.session_state.clip_val) if st.session_state.clip_val in num_options else 0, label_visibility="collapsed")
            with c5: 
                st.markdown("**BASE**")
                st.session_state.base_val = st.selectbox("BASE", num_options, index=num_options.index(st.session_state.base_val) if st.session_state.base_val in num_options else 0, label_visibility="collapsed")
            with c6: 
                st.markdown("**COVER**")
                st.session_state.cover_val = st.selectbox("COVER", num_options, index=num_options.index(st.session_state.cover_val) if st.session_state.cover_val in num_options else 0, label_visibility="collapsed")

            st.markdown("<br>**조립기**", unsafe_allow_html=True)
            render_grid_buttons(["1호기", "2호기", "3호기", "4호기", "5호기", "6호기"], "assembler_val", 6)

    elif step == 4:
        bad_qty = st.session_state.comp_def + st.session_state.front_def + st.session_state.rear_def + st.session_state.offset_def + st.session_state.etc_def
        total_qty = max(0, st.session_state.good_qty + bad_qty - st.session_state.shortage_qty)

        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 수량 등록</h4>", unsafe_allow_html=True)
            q1, q2, q3 = st.columns(3)
            with q1: 
                st.markdown("**검사 수량 (자동)**")
                st.text_input("검사 수량", value=f"{total_qty:,}", disabled=True, label_visibility="collapsed")
            with q2: 
                st.markdown("**양품수량**")
                if st.button(f"{st.session_state.good_qty:,}", key="f_good", use_container_width=True): 
                    val = str(st.session_state.good_qty)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("good_qty", "양품수량")
            with q3: 
                st.markdown("**불량수량 (자동)**")
                st.text_input("불량수량", value=f"{bad_qty:,}", disabled=True, label_visibility="collapsed")
        
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 불량 세부 내역</h4>", unsafe_allow_html=True)
            c1, c2, c3 = st.columns(3)
            with c1: 
                st.markdown("**완전불량**")
                if st.button(f"{st.session_state.comp_def:,}", key="f_comp", use_container_width=True): 
                    val = str(st.session_state.comp_def)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("comp_def", "완전불량")
            with c2: 
                st.markdown("**전면불량**")
                if st.button(f"{st.session_state.front_def:,}", key="f_front", use_container_width=True): 
                    val = str(st.session_state.front_def)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("front_def", "전면불량")
            with c3: 
                st.markdown("**배면불량**")
                if st.button(f"{st.session_state.rear_def:,}", key="f_rear", use_container_width=True): 
                    val = str(st.session_state.rear_def)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("rear_def", "배면불량")
            
            st.markdown("<br>", unsafe_allow_html=True)
            c4, c5, c6, c7 = st.columns(4)
            with c4: 
                st.markdown("**옵셋불량**")
                if st.button(f"{st.session_state.offset_def:,}", key="f_off", use_container_width=True): 
                    val = str(st.session_state.offset_def)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("offset_def", "옵셋불량")
            with c5: 
                st.markdown("**수량부족**")
                if st.button(f"{st.session_state.shortage_qty:,}", key="f_short", use_container_width=True): 
                    val = str(st.session_state.shortage_qty)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("shortage_qty", "수량부족")
            with c6: 
                st.markdown("**기타**")
                if st.button(f"{st.session_state.etc_def:,}", key="f_etc", use_container_width=True): 
                    val = str(st.session_state.etc_def)
                    st.session_state.numpad_buffer = val if val != "0" else ""
                    numpad_dialog("etc_def", "기타")
            with c7: 
                st.markdown("**OQC**")
                st.session_state.oqc_status = st.selectbox("OQC", ["선택안함", "육안", "OQC"], index=["선택안함", "육안", "OQC"].index(st.session_state.oqc_status), label_visibility="collapsed")

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

        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 실시간 수율 현황</h4>", unsafe_allow_html=True)
            rate_good = round((st.session_state.good_qty / total_qty) * 100, 1) if total_qty > 0 else 0.0
            
            c_yield, c_comp, c_front, c_rear, c_offset = "#10B981", "#EF4444", "#F59E0B", "#1e293b", "#8B5CF6"
            
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
            fig_bar.update_layout(barmode='overlay', showlegend=False, height=250, margin=dict(t=10, b=20, l=10, r=10), paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', xaxis=dict(showgrid=False, tickfont=dict(color='#1e293b')), yaxis=dict(showgrid=False, showticklabels=False, range=[0, y_max]))

            g_col1, g_col2 = st.columns(2)
            with g_col1: st.plotly_chart(fig_donut, use_container_width=True)
            with g_col2: st.plotly_chart(fig_bar, use_container_width=True)

        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 최종 확인 및 저장</h4>", unsafe_allow_html=True)
            rem_col, save_col = st.columns([0.7, 0.3])
            with rem_col:
                st.markdown("**비고**")
                st.session_state.remarks = st.text_area("비고", value=st.session_state.remarks, height=150, label_visibility="collapsed")
                
            with save_col:
                st.markdown("**&nbsp;**")
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
                                st.session_state.step = 5
                                st.session_state.comp_warned = st.session_state.front_warned = st.session_state.rear_warned = st.session_state.offset_warned = False
                                st.rerun()

    elif step == 5:
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 최근 저장 데이터 List</h4>", unsafe_allow_html=True)
            df_history = load_data().copy()
            if not df_history.empty:
                recent_10 = df_history.iloc[::-1].head(10).copy()
                st.dataframe(recent_10, use_container_width=True, hide_index=True)
            else:
                st.caption("저장된 데이터가 없습니다.")
