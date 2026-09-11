import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import json
import os
import re
from datetime import datetime, time as dt_time, timedelta, timezone
import time
from io import BytesIO
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

worker_list = ["작업자A", "작업자B", "작업자C", "작업자D", "작업자E", "작업자F"]
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
if "app_mode" not in st.session_state: st.session_state.app_mode = "START" 
if "step" not in st.session_state: st.session_state.step = 1
if "unlocked" in st.query_params:
    st.session_state.unlocked = True
    st.query_params.clear()

if "admin_authenticated" not in st.session_state: st.session_state.admin_authenticated = False

default_state = {
    "unique_id": "", "work_date": datetime.now(timezone(timedelta(hours=9))).date(), 
    "shift_type": "주간", "worker": "작업자A",
    "model_name": "D65S(KRIOS)", "lot_input_field": "", "in_date_field": datetime.now(timezone(timedelta(hours=9))).date(),
    "plating_type": "A", "start_date": datetime.now(timezone(timedelta(hours=9))).date(), "start_time": datetime.now(timezone(timedelta(hours=9))).time(),
    "end_date": datetime.now(timezone(timedelta(hours=9))).date(), "end_time": datetime.now(timezone(timedelta(hours=9))).time(), "unit": "1호기",
    "category": "1차 검사", "idle_time": 0, "painting_date": datetime.now(timezone(timedelta(hours=9))).date(),
    "painting_order": "", "painting_line": "A Line", 
    "clip_val": "1", "clip_k": False,
    "base_val": "1", "base_k": False,
    "cover_val": "1", "cover_k": False,
    "assembler_val": "1호기",
    "good_qty": 0, "comp_def": 0, "front_def": 0, "rear_def": 0, "offset_def": 0,
    "shortage_qty": 0, "etc_def": 0, "oqc_status": "선택안함", "remarks": "",
    "scanned_raw_data": "", "comp_warned": False, "front_warned": False, 
    "rear_warned": False, "offset_warned": False,
    "numpad_buffer": "", "timepad_buffer": "", "target_unique_id": ""
}

for key, value in default_state.items():
    if key not in st.session_state:
        st.session_state[key] = value

if not st.session_state.unlocked:
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
        [data-testid="stToolbar"] { display: none !important; }
        footer { display: none !important; }
        [data-testid="stAppViewContainer"] { background-color: #ffffff !important; }
    </style>
    """, unsafe_allow_html=True)
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        logo_l_data = get_image_base64("logo")
        if logo_l_data:
            st.markdown(f"<div style='text-align: center;'><img src='{logo_l_data}' style='max-width: 100%; max-height: 360px; object-fit: contain; margin-bottom: 20px;'></div>", unsafe_allow_html=True)
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
    st.markdown("<div style='position: fixed; bottom: 10%; left: 0; width: 100%; text-align: center; font-size: 10pt; color: #FFC000 !important; font-weight: bold;'>Created by --- Romero.K</div>", unsafe_allow_html=True)
    st.stop()


# ==============================================================================
# 💡 페이지별 CSS 분리 적용
# ==============================================================================
input_theme_css = """
<style>
footer { display: none !important; } 
[data-testid="collapsedControl"] { display: flex !important; visibility: visible !important; opacity: 1 !important; z-index: 99999 !important; }
body { overscroll-behavior-y: none !important; } 
::-webkit-scrollbar { display: none; }
.block-container { padding-top: 3rem !important; padding-bottom: 2rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 95% !important; }

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
    background-color: #E7E6E6 !important; 
    color: #000000 !important;
    border: 1px solid #cbd5e1 !important;
    box-shadow: none !important;
    transition: all 0.2s ease;
}
div[data-testid="stButton"] button:hover, div[data-testid="stButton"] button:focus, div[data-testid="stButton"] button:active {
    background-color: #1e293b !important;
    color: #ffffff !important;
    border-color: #1e293b !important;
}
div[data-testid="stButton"] button[kind="primary"] {
    background-color: #1e293b !important;
    color: #ffffff !important;
    border: 1px solid #0f172a !important;
    box-shadow: none !important;
}
div[data-testid="stButton"] button[kind="primary"]:hover { background-color: #0f172a !important; }

div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div {
    min-height: 2.6rem !important;
    border-radius: 8px !important;
    background-color: #E7E6E6 !important;
    border: 1px solid #cbd5e1 !important;
    padding: 0 !important;
    margin: 0 !important;
    box-shadow: none !important;
    box-sizing: border-box !important;
    transition: all 0.2s ease;
}
span[data-baseweb="tag"] { background-color: #1e293b !important; color: #ffffff !important; }
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div > div,
div[data-testid="stDateInput"] input,
div[data-testid="stTextInput"] input {
    min-height: 2.6rem !important;
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
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div > div:last-child { display: flex !important; align-items: center !important; }
div[data-baseweb="textarea"] textarea { 
    font-size: 1.1rem !important; 
    height: 70px !important;
    min-height: 70px !important; 
    max-height: 70px !important; 
    background-color: #E7E6E6 !important; 
    color: #000000 !important;
    border: 1px solid #cbd5e1 !important; 
    border-radius: 8px !important; 
    padding: 15px !important;
    transition: all 0.2s ease;
}
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div:focus-within,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div:focus-within {
    background-color: #1e293b !important;
    border-color: #1e293b !important;
}
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:focus-within > div,
div[data-testid="stDateInput"] div[data-baseweb="input"] > div:focus-within input,
div[data-testid="stTextInput"] div[data-baseweb="input"] > div:focus-within input { color: #ffffff !important; }
div[data-baseweb="textarea"]:focus-within textarea { background-color: #1e293b !important; color: #ffffff !important; border-color: #1e293b !important; }
div[data-baseweb="select"] input, div[data-baseweb="datepicker"] input { caret-color: transparent !important; cursor: pointer !important; }
input[placeholder*="SCAN APP"] { color: #000000 !important; font-weight: 900 !important; }
input[placeholder*="SCAN APP"]::placeholder { color: #4b5563 !important; font-weight: bold !important; opacity: 0.8 !important; }

[data-testid="stSidebar"] { background-color: #0f172a !important; }
[data-testid="stSidebar"] * { color: #f8fafc !important; }
[data-testid="stSidebar"] .stButton > button { 
    height: 48px !important; 
    max-height: 48px !important;
    justify-content: flex-start !important; 
    padding-left: 15px !important; 
    margin-bottom: 5px !important; 
    border-radius: 6px !important; 
    background-color: transparent !important;
    border: 1px solid transparent !important; 
    color: #8B9CB6 !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button p { font-weight: 800 !important; font-size: 14px !important; text-indent: 10px !important; text-align: left !important; }
[data-testid="stSidebar"] .stButton > button[kind="primary"] { background-color: #1e293b !important; color: #FFFFFF !important; border: none !important; border-left: 4px solid #FFC000 !important; }
[data-testid="stSidebar"] .stButton > button[kind="secondary"] { background-color: transparent !important; color: #8B9CB6 !important; border: 1px solid transparent !important; }
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover { background-color: #1e293b !important; color: #ffffff !important; border: 1px solid transparent !important; }
div[data-testid="stCheckbox"] { display: flex; align-items: center; height: 2.6rem; padding-left: 10px; }
</style>
"""

analysis_theme_css = """
<style>
footer { display: none !important; } 
[data-testid="collapsedControl"] { display: flex !important; visibility: visible !important; opacity: 1 !important; z-index: 99999 !important; }
body { overscroll-behavior-y: none !important; } 
::-webkit-scrollbar { display: none; }
.block-container { padding-top: 2.5rem !important; padding-bottom: 2rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 98% !important; }

/* Dark Background */
[data-testid="stAppViewContainer"] { background-color: #050814 !important; color: #E2E8F0 !important; font-family: 'Consolas', 'Courier New', monospace !important; }
[data-testid="stSidebar"] { background-color: #030508 !important; }
[data-testid="stSidebar"] * { color: #94A3B8 !important; font-family: 'Consolas', 'Courier New', monospace !important; }

/* Dashboard Cards */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background-color: #0B101E !important;
    border: 1px solid #1E293B !important;
    border-radius: 6px !important;
    box-shadow: 0 0 15px rgba(0, 0, 0, 0.6) !important;
    padding: 1.2rem !important;
    margin-bottom: 0.8rem !important;
}

h1, h2, h3, h4, h5, h6, p, span, div { color: #94A3B8 !important; font-family: 'Consolas', 'Courier New', monospace !important; }

/* Custom Neon Titles */
.command-header { color: #38BDF8 !important; font-weight: 900 !important; letter-spacing: 1.5px; text-shadow: 0 0 5px rgba(56, 189, 248, 0.4); }
.metric-label { color: #38BDF8 !important; font-size: 1rem !important; font-weight: bold !important; letter-spacing: 1px; margin-bottom: 5px; }

/* Live Blinking Dot Effect */
@keyframes blink {
    0% { opacity: 1; box-shadow: 0 0 12px #EF4444; }
    50% { opacity: 0.3; box-shadow: 0 0 3px #EF4444; }
    100% { opacity: 1; box-shadow: 0 0 12px #EF4444; }
}
.live-dot {
    height: 14px; width: 14px; background-color: #EF4444; border-radius: 50%;
    display: inline-block; margin-right: 12px; margin-bottom: 2px;
    animation: blink 1.2s ease-in-out infinite;
}

/* Custom Selectbox for Dark Mode */
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
    background-color: #0F172A !important;
    border: 1px solid #1E293B !important;
    color: #E2E8F0 !important;
}
div[data-testid="stSelectbox"] div[data-baseweb="select"] > div > div {
    color: #E2E8F0 !important;
    font-weight: bold !important;
    font-size: 1.1rem !important;
}

/* Buttons in Dark Mode */
div[data-testid="stButton"] button {
    background-color: #0F172A !important;
    color: #38BDF8 !important;
    border: 1px solid #1E293B !important;
    border-radius: 4px !important;
    font-weight: bold !important;
    transition: all 0.2s ease;
}
div[data-testid="stButton"] button:hover {
    background-color: #1E293B !important;
    border-color: #38BDF8 !important;
    box-shadow: 0 0 10px rgba(56, 189, 248, 0.3) !important;
}

/* Sidebar Custom Buttons */
[data-testid="stSidebar"] .stButton > button { 
    height: 48px !important; max-height: 48px !important; justify-content: flex-start !important; padding-left: 15px !important; margin-bottom: 5px !important; border-radius: 4px !important; background-color: transparent !important; border: 1px solid transparent !important; box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] { background-color: #0F172A !important; color: #38BDF8 !important; border-left: 4px solid #38BDF8 !important; }
[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover { background-color: #0F172A !important; color: #E2E8F0 !important; }
</style>
"""

if st.session_state.current_page == "input":
    st.markdown(input_theme_css, unsafe_allow_html=True)
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
                    if (text.includes('⬅️ 이전') || text.includes('다음 ➡️') || text.includes('Data 최종 저장') || text.includes('작업시작 등록') || text.trim() === '적용' || text.includes('신규 작업 등록')) { 
                        btn.style.setProperty('background', '#305496', 'important');
                        btn.style.setProperty('background-color', '#305496', 'important');
                        btn.style.setProperty('border', '1px solid #203864', 'important');
                        btn.style.setProperty('color', '#ffffff', 'important');
                        btn.style.setProperty('box-shadow', 'none', 'important');
                    }
                    if (text.includes('⬅️ 이전') || text.includes('다음 ➡️') || text.includes('신규 작업 등록') || text.includes('작업시작 등록') || text.includes('Data 최종 저장')) {
                        btn.style.setProperty('height', '70px', 'important'); 
                        btn.style.setProperty('max-height', '70px', 'important');
                        btn.style.setProperty('font-size', '1.2rem', 'important');
                        btn.style.setProperty('margin-top', '0px', 'important');
                    }
                    if (text.trim() === 'ADMINISTRATOR') { 
                        btn.style.setProperty('background', '#1e293b', 'important');
                        btn.style.setProperty('background-color', '#1e293b', 'important');
                        btn.style.setProperty('color', '#FFC000', 'important');
                        btn.style.setProperty('border', '1px solid #0f172a', 'important');
                        btn.style.setProperty('height', '7.2rem', 'important');
                        btn.style.setProperty('min-height', '7.2rem', 'important');
                        btn.style.setProperty('max-height', '7.2rem', 'important');
                        btn.style.setProperty('font-size', '1.4rem', 'important');
                        btn.style.setProperty('font-weight', '900', 'important');
                    }
                });
            };
            const styleScanner = () => {
                if (!window.parent.document) return;
                window.parent.document.querySelectorAll('input').forEach(el => {
                    if (el.getAttribute('placeholder') && el.getAttribute('placeholder').includes('SCAN APP')) {
                        el.style.setProperty('font-size', '1.2rem', 'important');
                        el.style.setProperty('font-weight', '900', 'important');
                        let parentDiv = el.parentElement;
                        let grandParent = el.closest('div[data-baseweb="input"]');
                        if (parentDiv) parentDiv.style.setProperty('border', 'none', 'important');
                        if (grandParent) grandParent.style.setProperty('border', '2px solid #eab308', 'important');
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
                        el.addEventListener('focus', function(e) { e.target.blur(); });
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
else:
    st.markdown(analysis_theme_css, unsafe_allow_html=True)

def render_grid_buttons(options, state_key, columns, use_width=True):
    rows = [options[i:i+columns] for i in range(0, len(options), columns)]
    for row_opts in rows:
        cols = st.columns(columns)
        for i, opt in enumerate(row_opts):
            with cols[i]:
                if opt.strip() == "": st.write("") 
                else:
                    btn_type = "primary" if st.session_state.get(state_key) == opt else "secondary"
                    if st.button(opt, key=f"btn_{state_key}_{opt}", type=btn_type, use_container_width=use_width):
                        st.session_state[state_key] = opt
                        st.rerun()

SCOPE = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
EXCEL_COLUMNS = [
    "고유 ID", "상태", "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", 
    "모델명(MI)", "도금구분", "UPH", "UPD", "검사 수량", "양품수량", "양품 수량(전/배 포함)", 
    "불량수량", "양품율", "양품율(전/배 포함)", "완전불량율", "전면불량율", "배면불량율", 
    "완전불량", "전면불량", "배면불량", "옵셋불량", "수량부족", "기타", "OQC", "비고", 
    "도장라인", "도장일", "도장순서", "입고일", "LOT NO.", "CLIP", "BASE", "COVER", 
    "조립기", "월", "작업자"
]
SPREADSHEET_ID = "1DeMJJkuq7bYa4XNK_NbkqZ-vOJKqGhmYXIvHm3yJl8E"
# 💡 핵심 수정: 사용자가 지정한 시트명으로 완전히 고정
TAB_NAME = "2026년 3Q"

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

def get_sheet():
    doc = get_spreadsheet_doc()
    if doc:
        try: 
            return doc.worksheet(TAB_NAME)
        except Exception as e: 
            st.error(f"🚨 '{TAB_NAME}' 시트 탭을 찾을 수 없습니다: {e}")
            return doc.sheet1
    return None

@st.cache_data(ttl=15) 
def load_analysis_data():
    doc = get_spreadsheet_doc()
    if doc is None: return pd.DataFrame()
    
    try:
        ws = doc.worksheet(TAB_NAME)
        raw_data = ws.get_all_values()
    except Exception as e:
        st.error(f"🚨 데이터를 불러오는 중 오류 발생: {e}")
        return pd.DataFrame()
        
    if len(raw_data) < 2: return pd.DataFrame()
    
    header_idx = -1
    for i, row in enumerate(raw_data[:20]):
        row_str = "".join(str(c).replace(" ", "") for c in row)
        if any(k in row_str for k in ["날짜", "일자", "모델", "고유ID", "LOT"]):
            header_idx = i; break
            
    if header_idx == -1: return pd.DataFrame()
    
    headers = [str(h).strip() for h in raw_data[header_idx]]
    df_sheet = pd.DataFrame(raw_data[header_idx+1:], columns=headers)
    
    rename_dict = {}
    for c in df_sheet.columns:
        cc = str(c).replace(" ", "").replace("률", "율").upper()
        if "모델" in cc: rename_dict[c] = '모델명(MI)'
        elif "날짜" in cc or "일자" in cc: rename_dict[c] = '날짜'
        elif "시작" in cc and "시간" in cc: rename_dict[c] = '시작시간'
        elif "소요" in cc and "시간" in cc: rename_dict[c] = '소요시간'
        elif "LOT" in cc: rename_dict[c] = 'LOT NO.'
        elif "구분" in cc: rename_dict[c] = '구분'
        elif "도금" in cc: rename_dict[c] = '도금구분'
        elif cc == "양품율" or cc == "1차양품율": rename_dict[c] = '양품율'
        elif "전" in cc and "배" in cc and "양품율" in cc: rename_dict[c] = '양품율(전/배 포함)'
        elif "완전" in cc and "불량" in cc: rename_dict[c] = '완전불량율'
        elif "전면" in cc and "불량" in cc: rename_dict[c] = '전면불량율'
        elif "배면" in cc and "불량" in cc: rename_dict[c] = '배면불량율'
    
    df_sheet = df_sheet.rename(columns=rename_dict)
    return df_sheet

@st.cache_data(ttl=60)
def load_data():
    sheet = get_sheet()
    if sheet is None: return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row'])
    try:
        raw_data = sheet.get_all_values()
        if len(raw_data) < 2: return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row'])
        
        header_idx = -1
        for i, row in enumerate(raw_data[:15]):
            row_str = "".join(str(c).replace(" ", "") for c in row)
            if "날짜" in row_str or "교대" in row_str or "고유ID" in row_str.upper():
                header_idx = i; break
                
        if header_idx == -1: return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row'])
        
        headers = [str(h).strip() for h in raw_data[header_idx]]
        clean_headers = {str(c).replace(" ", "").replace("률", "율").upper(): c for c in headers}
        
        EXCEL_COLS_LOAD = ["고유 ID", "상태", "날짜", "교대", "시작시간", "종료시간", "휴동시간", "소요시간", "구분", "호기", "모델명(MI)", "검사 수량", "양품수량", "불량수량", "LOT NO."]
        
        data_list = []
        for r_idx in range(header_idx + 1, len(raw_data)):
            row = raw_data[r_idx]
            if any(str(c).strip() for c in row):
                row_data = {"_sheet_row": r_idx + 1}
                for col in EXCEL_COLS_LOAD:
                    col_key = col.replace(" ", "").replace("률", "율").upper()
                    if col_key in clean_headers:
                        try:
                            c_idx = headers.index(clean_headers[col_key])
                            row_data[col] = row[c_idx] if c_idx < len(row) else ""
                        except: row_data[col] = ""
                    else: row_data[col] = ""
                data_list.append(row_data)
        
        result_df = pd.DataFrame(data_list)
        return result_df
    except: return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row'])

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
        st.cache_data.clear() 
        return True
    except Exception as e:
        st.error(f"데이터 저장 오류: {e}")
        return False

def parse_scanned_data():
    raw_val = st.session_state.get("scanned_raw_data", "")
    if not raw_val: return
    st.session_state.unique_id = raw_val
    if '$' in raw_val:
        parts = [p for p in raw_val.split('$') if p]
        if len(parts) >= 5:
            st.session_state.lot_input_field = parts[4]
        else:
            st.session_state.lot_input_field = parts[-1]
    else:
        st.session_state.lot_input_field = raw_val
    st.session_state.scanned_raw_data = "" 

def on_scan_apply(): parse_scanned_data()

def pad_callback(digit):
    c_val = st.session_state.get("numpad_buffer", "")
    if digit == "C": st.session_state.numpad_buffer = ""
    elif digit == "⬅": st.session_state.numpad_buffer = c_val[:-1]
    else:
        if len(c_val) < 8: st.session_state.numpad_buffer = c_val + digit

@st.dialog("🔢 수량 입력 패드")
def numpad_dialog(field_key, display_name):
    c_val = st.session_state.get("numpad_buffer", "")
    st.markdown(f"<div style='text-align:center; font-size:1.8rem; font-weight:bold; color:#1e293b; padding:15px; background:#f8fafc; border-radius:10px; margin-bottom:15px; border:1px solid #cbd5e1;'>{display_name}<br><span style='color:#1e293b; font-size:2.5rem;'>{c_val if c_val else '0'}</span></div>", unsafe_allow_html=True)
    pad_rows = [["7", "8", "9"], ["4", "5", "6"], ["1", "2", "3"], ["C", "0", "⬅"]]
    for r in pad_rows:
        cols = st.columns(3)
        for i, val in enumerate(r):
            with cols[i]:
                st.button(val, key=f"pad_{field_key}_{val}", use_container_width=True, on_click=pad_callback, args=(val,))
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("적용 (Enter)", type="primary", use_container_width=True):
        if field_key in ["painting_order", "clip_val", "base_val", "cover_val"]:
            st.session_state[field_key] = str(st.session_state.get("numpad_buffer", ""))
        else:
            val = st.session_state.get("numpad_buffer", "")
            st.session_state[field_key] = int(val) if val else 0
        st.session_state.numpad_buffer = "" 
        st.rerun()

def timepad_callback(digit):
    c_val = st.session_state.get("timepad_buffer", "")
    if digit == "C": st.session_state.timepad_buffer = ""
    elif digit == "⬅": st.session_state.timepad_buffer = c_val[:-1]
    else:
        if len(c_val) < 4: st.session_state.timepad_buffer = c_val + digit

@st.dialog("⏰ 시간 입력 패드 (HH:MM)")
def timepad_dialog(field_key, display_name):
    c_val = st.session_state.get("timepad_buffer", "")
    display_str = c_val.ljust(4, "_")
    display_str = f"{display_str[:2]}:{display_str[2:]}"
    st.markdown(f"<div style='text-align:center; font-size:1.5rem; font-weight:bold; color:#1e293b; padding:15px; background:#f8fafc; border-radius:10px; margin-bottom:15px; border:1px solid #cbd5e1;'>{display_name}<br><span style='color:#1e293b; font-size:2.5rem; letter-spacing: 2px;'>{display_str}</span></div>", unsafe_allow_html=True)
    pad_rows = [["7", "8", "9"], ["4", "5", "6"], ["1", "2", "3"], ["C", "0", "⬅"]]
    for r in pad_rows:
        cols = st.columns(3)
        for i, val in enumerate(r):
            with cols[i]:
                st.button(val, key=f"tpad_{field_key}_{val}", use_container_width=True, on_click=timepad_callback, args=(val,))
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("적용 (Enter)", type="primary", use_container_width=True):
        buffer_val = st.session_state.get("timepad_buffer", "")
        if len(buffer_val) == 4:
            try:
                h = int(buffer_val[:2])
                m = int(buffer_val[2:])
                if 0 <= h <= 23 and 0 <= m <= 59:
                    st.session_state[field_key] = dt_time(h, m)
                    st.session_state.timepad_buffer = "" 
                    st.rerun()
            except ValueError: pass

@st.dialog("SBL Warning!")
def show_sbl_warning(defect_type, rate):
    st.markdown(f"### [{defect_type}] 불량 제품 별도 보관 조치")
    st.error(f"현재 1차검사 공정의 {defect_type}율이 **{rate:.1f}%** 로 기준치(5.0%)를 초과하였습니다.")
    if st.button("확인 완료 (닫기)", key=f"btn_close_{defect_type}"):
        st.rerun()

def render_nav_buttons(step_num, max_step):
    st.markdown("<br>", unsafe_allow_html=True)
    c_nav = st.columns(6)
    with c_nav[4]:
        if step_num > 1:
            if st.button("⬅️ 이전", use_container_width=True):
                st.session_state.step -= 1
                st.rerun()
    with c_nav[5]:
        if step_num < max_step:
            if st.button("다음 ➡️", use_container_width=True):
                st.session_state.step += 1
                st.rerun()

@st.dialog("🔒 관리자 인증")
def admin_auth_dialog():
    st.markdown("<div style='color:#94A3B8; margin-bottom:10px;'>분석 데이터를 확인하려면 관리자 비밀번호를 입력하세요.</div>", unsafe_allow_html=True)
    pwd = st.text_input("비밀번호", type="password", label_visibility="collapsed", placeholder="비밀번호 입력")
    if st.button("✅ 확인", type="primary", use_container_width=True):
        if pwd == "6233":
            st.session_state.admin_authenticated = True
            st.rerun()
        else:
            st.error("비밀번호가 일치하지 않습니다.")

# ==========================================
# 💡 Administrator (LIVE YIELD COMMAND CENTER)
# ==========================================
if st.session_state.current_page == "analysis":
    if not st.session_state.admin_authenticated:
        admin_auth_dialog()
        st.stop()
        
    col1, col2, col3 = st.columns([0.7, 0.15, 0.15])
    with col1:
        st.markdown(f"<div class='command-header' style='font-size: 1.8rem; margin-top: 5px;'><span class='live-dot'></span>LIVE YIELD COMMAND CENTER v1.0</div>", unsafe_allow_html=True)
        st.markdown("<div style='color: #64748B; font-size: 0.85rem; margin-bottom: 15px;'>Manual LOT surveillance pipeline active.</div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 REFRESH DATA", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
    with col3:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("RETURN TO INPUT", type="primary", use_container_width=True):
            st.session_state.current_page = "input"
            st.session_state.admin_authenticated = False 
            st.rerun()
            
    df = load_analysis_data().copy()

    if df.empty or '모델명(MI)' not in df.columns: 
        st.warning("데이터베이스에 유효한 정보가 없습니다.")
    else:
        def pct_to_float(x):
            try:
                if pd.isna(x) or str(x).strip() == '': return np.nan
                return float(str(x).replace('%', '').replace(',', '').strip())
            except: return np.nan

        df['Yield_1'] = df.get('양품율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
        df['Yield_2'] = df.get('양품율(전/배 포함)', pd.Series([np.nan]*len(df))).apply(pct_to_float)
        df['Def_Comp'] = df.get('완전불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
        df['Def_Front'] = df.get('전면불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
        df['Def_Rear'] = df.get('배면불량율', pd.Series([np.nan]*len(df))).apply(pct_to_float)
        
        # 💡 날짜 파서 (에러 방어용)
        def parse_dt(r):
            try:
                d_val = r.get('날짜', '')
                t_val = r.get('시작시간', '00:00')
                if pd.isna(d_val) or str(d_val).strip() == '': return pd.NaT
                if pd.isna(t_val) or str(t_val).strip() == '': t_val = "00:00"
                t_clean = re.sub(r'[^\d]', '', str(t_val))
                if len(t_clean) >= 4: t_str = f"{t_clean[:2]}:{t_clean[2:4]}"
                elif len(t_clean) == 3: t_str = f"0{t_clean[:1]}:{t_clean[1:3]}"
                elif len(t_clean) in [1, 2]: t_str = f"{t_clean.zfill(2)}:00"
                else: t_str = "00:00"
                d_str_val = str(d_val).strip()
                m = re.search(r'(\d{4})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{1,2})', d_str_val)
                if m: return pd.to_datetime(f"{m.group(1)}-{m.group(2)}-{m.group(3)} {t_str}")
                m_ddmmyy = re.search(r'^(\d{2})\s*-\s*(\d{2})\s*-\s*(\d{4})$', d_str_val)
                if m_ddmmyy: return pd.to_datetime(f"{m_ddmmyy.group(3)}-{m_ddmmyy.group(2)}-{m_ddmmyy.group(1)} {t_str}")
            except: pass
            return pd.NaT
            
        df['DateTime'] = df.apply(parse_dt, axis=1)
        
        # 💡 1차 검사 강제 필터링
        if '구분' in df.columns:
            df = df[df['구분'].fillna('').astype(str).str.contains('1차', na=False)]
            
        # 💡 도금구분 정리
        if '도금구분' in df.columns:
            df['도금구분'] = df['도금구분'].fillna('A').astype(str).str.strip().str.upper()
            df['도금구분'] = df['도금구분'].apply(lambda x: 'B' if 'B' in x else 'A')
            
        # 💡 강제 출력 스캔 (가장 최신 150개)
        df_target = df.tail(150).copy()
        
        with st.container(border=True):
            col_a, col_b = st.columns(2)
            with col_a:
                st.markdown("<div class='metric-label'>■ TARGET MODEL SELECTION (Multi)</div>", unsafe_allow_html=True)
                models_available = sorted(df_target['모델명(MI)'].replace('', np.nan).dropna().unique().tolist()) if '모델명(MI)' in df_target.columns else []
                selected_models = st.multiselect("Select Models", models_available, default=models_available[:1] if models_available else [], label_visibility="collapsed")
            
            with col_b:
                st.markdown("<div class='metric-label'>■ PLATING TYPE (도금구분)</div>", unsafe_allow_html=True)
                plating_available = sorted(df_target['도금구분'].unique().tolist()) if '도금구분' in df_target.columns else ['A', 'B']
                selected_plating = st.multiselect("Plating Type", plating_available, default=plating_available, label_visibility="collapsed")

            if not selected_models:
                st.warning("Please select at least one model to render charts.")
            else:
                model_df = df_target[df_target['모델명(MI)'].isin(selected_models)].copy()
                if selected_plating and not model_df.empty and '도금구분' in model_df.columns:
                    model_df = model_df[model_df['도금구분'].isin(selected_plating)]
                
                if model_df.empty:
                    st.info("No data available for the selected model & plating type.")
                else:
                    # 💡 X축을 실제 시간으로 매핑
                    model_df['DateTime'] = model_df['DateTime'].fillna(pd.Timestamp('1900-01-01'))
                    model_df = model_df.sort_values(['DateTime'])
                    
                    model_df['LOT NO.'] = model_df.get('LOT NO.', pd.Series(['UNKNOWN']*len(model_df)))
                    model_df['LOT NO.'] = model_df['LOT NO.'].replace({'': 'UNKNOWN', 'nan': 'UNKNOWN', None: 'UNKNOWN'}).fillna('UNKNOWN').astype(str)
                    
                    def make_hover_text(row):
                        mod_str = str(row.get('모델명(MI)', ''))
                        plat_str = str(row.get('도금구분', 'A'))
                        dur_str = str(row.get('소요시간', '0'))
                        lot_str = str(row['LOT NO.'])
                        return f"[{mod_str} - {plat_str}]<br>LOT: {lot_str}<br>소요: {dur_str}분"
                    
                    model_df['HoverText'] = model_df.apply(make_hover_text, axis=1)

                    def get_dark_layout(title_text, y_title):
                        return dict(
                            title=dict(text=f"■ {title_text}", font=dict(color='#E2E8F0', size=16)),
                            plot_bgcolor='#0B101E', paper_bgcolor='#0B101E',
                            font=dict(color='#94A3B8', family='monospace'),
                            xaxis=dict(title='작업 시작시간', type='date', showgrid=True, gridcolor='#1E293B', linecolor='#334155', tickformat='%m-%d %H:%M'),
                            yaxis=dict(title=y_title, showgrid=True, gridcolor='#1E293B', linecolor='#334155', zeroline=False),
                            margin=dict(l=40, r=40, t=50, b=40), 
                            hovermode='x unified',
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(color='#E2E8F0'))
                        )

                    with st.container(border=True):
                        fig1 = go.Figure()
                        colors_1 = ['#00E5FF', '#FF9900', '#00FF00', '#FFFF00']
                        colors_2 = ['#FF00FF', '#FF3366', '#9D00FF', '#00BFFF']
                        
                        for idx, mod in enumerate(selected_models):
                            m_df = model_df[model_df['모델명(MI)'] == mod].dropna(subset=['Yield_1'])
                            if m_df.empty: continue
                            
                            c1 = colors_1[idx % len(colors_1)]
                            c2 = colors_2[idx % len(colors_2)]
                            
                            fig1.add_trace(go.Scatter(
                                x=m_df['DateTime'], y=m_df['Yield_1'], name=f"[{mod}] 1차", 
                                mode='lines+markers+text', text=m_df['LOT NO.'], textposition='top center',
                                textfont=dict(size=10, color=c1),
                                line=dict(color=c1, width=2), marker=dict(size=6, color=c1), hovertext=m_df['HoverText']
                            ))
                            fig1.add_trace(go.Scatter(
                                x=m_df['DateTime'], y=m_df['Yield_2'], name=f"[{mod}] 전/배포함", 
                                mode='lines+markers', line=dict(color=c2, width=2, dash='dot'), 
                                marker=dict(size=5, symbol='x', color=c2), hovertext=m_df['HoverText']
                            ))
                        
                        fig1.update_layout(**get_dark_layout("1ST YIELD TREND (STANDARD vs INCL. F/R)", "YIELD (%)"), height=400)
                        st.plotly_chart(fig1, use_container_width=True)

                    with st.container(border=True):
                        fig2 = go.Figure()
                        for idx, mod in enumerate(selected_models):
                            m_df = model_df[model_df['모델명(MI)'] == mod].dropna(subset=['Def_Comp'])
                            if m_df.empty: continue
                            c1 = colors_2[idx % len(colors_2)]
                            
                            fig2.add_trace(go.Scatter(
                                x=m_df['DateTime'], y=m_df['Def_Comp'], name=f"[{mod}] 완전불량", 
                                mode='lines+markers+text', text=m_df['LOT NO.'], textposition='top center',
                                textfont=dict(size=10, color=c1),
                                line=dict(color=c1, width=2), fill='tozeroy', fillcolor=c1.replace(')', ', 0.1)').replace('rgb', 'rgba') if 'rgb' in c1 else None,
                                marker=dict(size=6, color=c1), hovertext=m_df['HoverText']
                            ))
                            
                        fig2.update_layout(**get_dark_layout("COMPLETE DEFECT RATE", "DEFECT RATE (%)"), height=350)
                        st.plotly_chart(fig2, use_container_width=True)

                    with st.container(border=True):
                        fig3 = go.Figure()
                        for idx, mod in enumerate(selected_models):
                            m_df = model_df[model_df['모델명(MI)'] == mod].dropna(subset=['Def_Front', 'Def_Rear'], how='all')
                            if m_df.empty: continue
                            c1 = colors_1[idx % len(colors_1)]
                            c2 = colors_2[idx % len(colors_2)]
                            
                            fig3.add_trace(go.Scatter(
                                x=m_df['DateTime'], y=m_df['Def_Front'], name=f"[{mod}] 전면불량", 
                                mode='lines+markers+text', text=m_df['LOT NO.'], textposition='top center',
                                textfont=dict(size=10, color=c1),
                                line=dict(color=c1, width=2), marker=dict(size=6, color=c1), hovertext=m_df['HoverText']
                            ))
                            fig3.add_trace(go.Scatter(
                                x=m_df['DateTime'], y=m_df['Def_Rear'], name=f"[{mod}] 배면불량", 
                                mode='lines+markers', line=dict(color=c2, width=2, dash='dash'), 
                                marker=dict(size=6, symbol='triangle-up', color=c2), hovertext=m_df['HoverText']
                            ))
                            
                        fig3.update_layout(**get_dark_layout("FRONT & REAR DEFECT RATE", "DEFECT RATE (%)"), height=350)
                        st.plotly_chart(fig3, use_container_width=True)

# ==========================================
# Main Input App
# ==========================================
elif st.session_state.current_page == "input":
    
    top_c1, top_c2 = st.columns([5, 1])
    with top_c1:
        logo_s_data = get_image_base64("at")
        img_html = f"<img src='{logo_s_data}' style='height: 96px; margin-right: 20px;'>" if logo_s_data else ""
        st.markdown(
            f"<div style='background: #ffffff; padding: 0 30px; border-radius: 8px; margin-bottom: 15px; border: 1px solid #cbd5e1; height: 7.2rem; display: flex; align-items: center; justify-content: center; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04); box-sizing: border-box;'>"
            f"{img_html}"
            f"<h3 style='color: #1e293b; margin: 0; font-weight: 900; font-size: 1.8rem; letter-spacing: 1px;'>VISION DATA KEY-IN SYSTEM</h3>"
            f"</div>", 
            unsafe_allow_html=True
        )
    with top_c2:
        if st.button("ADMINISTRATOR", use_container_width=True, type="primary"):
            st.session_state.current_page = "analysis"
            st.rerun()

    with st.sidebar:
        KST = timezone(timedelta(hours=9))
        now = datetime.now(KST)
        weekdays = ['월', '화', '수', '목', '금', '토', '일']
        current_time_str = f"{now.strftime('%Y년 %m월 %d일')} ({weekdays[now.weekday()]}) {now.strftime('%p %I:%M').replace('AM', '오전').replace('PM', '오후')}"
        st.markdown(f"<div style='text-align: center; color: #000000 !important; background-color: #f1f5f9 !important; padding: 10px; border-radius: 8px; font-weight: bold; font-size: 0.85rem; margin-bottom: 20px;'>{current_time_str}</div>", unsafe_allow_html=True)
        
        st.markdown("<h4 style='color: #f8fafc; font-size: 1.1rem; border-bottom: 1px solid #334155; padding-bottom: 8px;'>■ 시작 프로세스</h4><br>", unsafe_allow_html=True)
        if st.button("작업 등록", type="primary" if (st.session_state.app_mode=="START" and st.session_state.step==1) else "secondary", use_container_width=True):
            st.session_state.app_mode = "START"
            st.session_state.step = 1
            st.rerun()
        if st.button("생산 시작", type="primary" if (st.session_state.app_mode=="START" and st.session_state.step==2) else "secondary", use_container_width=True):
            st.session_state.app_mode = "START"
            st.session_state.step = 2
            st.rerun()
        if st.button("도장 및 조립", type="primary" if (st.session_state.app_mode=="START" and st.session_state.step==3) else "secondary", use_container_width=True):
            st.session_state.app_mode = "START"
            st.session_state.step = 3
            st.rerun()

        st.markdown("<br><h4 style='color: #f8fafc; font-size: 1.1rem; border-bottom: 1px solid #334155; padding-bottom: 8px;'>■ 종료 프로세스</h4><br>", unsafe_allow_html=True)
        if st.button("작업 종료", type="primary" if (st.session_state.app_mode=="END" and st.session_state.step==1) else "secondary", use_container_width=True):
            st.session_state.app_mode = "END"
            st.session_state.step = 1
            st.rerun()
        if st.button("VISION DATA", type="primary" if (st.session_state.app_mode=="END" and st.session_state.step==2) else "secondary", use_container_width=True):
            st.session_state.app_mode = "END"
            st.session_state.step = 2
            st.rerun()

        st.markdown("<br><h4 style='color: #f8fafc; font-size: 1.1rem; border-bottom: 1px solid #334155; padding-bottom: 8px;'>■ 데이터 관리</h4><br>", unsafe_allow_html=True)
        if st.button("최근 저장 Data List", type="primary" if st.session_state.app_mode=="EDIT" else "secondary", use_container_width=True):
            st.session_state.app_mode = "EDIT"
            st.session_state.step = 1
            st.rerun()
            
        st.markdown("<hr style='border-color: #334155; margin-top: 10px; margin-bottom: 10px;'>", unsafe_allow_html=True)
        st.markdown("<div style='text-align: center; color: #FFC000 !important; font-size: 14px; font-weight: bold;'>Created by --- Romero.K</div>", unsafe_allow_html=True)

    step = st.session_state.step

    # ==========================================
    # 💡 1. [작업 시작] 모드 
    # ==========================================
    if st.session_state.app_mode == "START":
        if step == 1:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 작업 정보</h4><br>", unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                with c1: 
                    st.markdown("**근무일자**")
                    w_date_val = st.session_state.get("work_date", datetime.now(timezone(timedelta(hours=9))).date())
                    st.session_state.work_date = st.date_input("근무일자", value=w_date_val, label_visibility="collapsed")
                with c2: 
                    st.markdown("**모델명**")
                    m_val = st.session_state.get("model_name", "D65S(KRIOS)")
                    st.session_state.model_name = st.selectbox("모델명", model_list, index=model_list.index(m_val) if m_val in model_list else 0, label_visibility="collapsed")
                with c3:
                    st.markdown("**교대**")
                    render_grid_buttons(["주간", "야간"], "shift_type", 2, use_width=True)
                with c4:
                    st.markdown("**작업자**")
                    w_val = st.session_state.get("worker", worker_list[0])
                    st.session_state.worker = st.selectbox("작업자", worker_list, index=worker_list.index(w_val) if w_val in worker_list else 0, label_visibility="collapsed")

            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ LOT 정보</h4><br>", unsafe_allow_html=True)
                sc1, sc2, sc3, sc4, sc5, sc6 = st.columns(6)
                with sc1:
                    st.markdown("**SCAN DATA**")
                    st.text_input("SCAN DATA", key="scanned_raw_data", label_visibility="collapsed", placeholder="SCAN APP")
                with sc2:
                    st.markdown("**고유 ID (자동할당)**")
                    st.text_input("고유 ID", value=st.session_state.get("unique_id", ""), disabled=True, label_visibility="collapsed")
                with sc3:
                    st.markdown("**&nbsp;**")
                    st.button("적용", type="primary", use_container_width=True, on_click=on_scan_apply)
                with sc4:
                    st.markdown("**LOT (적용됨)**")
                    st.text_input("LOT", value=st.session_state.get("lot_input_field", ""), disabled=True, label_visibility="collapsed")
                with sc5:
                    st.markdown("**입고일 (적용됨)**")
                    in_date_val = st.session_state.get("in_date_field", datetime.now(timezone(timedelta(hours=9))).date())
                    st.date_input("입고일", value=in_date_val, disabled=True, label_visibility="collapsed")
                with sc6:
                    st.markdown("**도금구분**")
                    render_grid_buttons(["A", "B"], "plating_type", 2, use_width=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            c_nav = st.columns(6)
            with c_nav[5]:
                if st.button("다음 ➡️", use_container_width=True):
                    st.session_state.step = 2
                    st.rerun()

        elif step == 2:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 시작 등록</h4><br>", unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                with c1: 
                    st.markdown("**시작일**")
                    s_date_val = st.session_state.get("start_date", datetime.now(timezone(timedelta(hours=9))).date())
                    st.session_state.start_date = st.date_input("시작일", value=s_date_val, label_visibility="collapsed")
                with c2: 
                    st.markdown("**시작시간**")
                    start_time_obj = st.session_state.get("start_time")
                    display_val = start_time_obj.strftime("%H:%M") if start_time_obj else "입력"
                    if st.button(display_val, key="btn_start_time", use_container_width=True):
                        st.session_state.timepad_buffer = ""
                        timepad_dialog("start_time", "시작시간")
                with c3: st.write("")
                with c4: st.write("")
            
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 설비 정보</h4><br>", unsafe_allow_html=True)
                render_grid_buttons(["1호기", "2호기", "3호기", "4호기", "5호기", "6호기"], "unit", 6)
            
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 검사 구분</h4><br>", unsafe_allow_html=True)
                render_grid_buttons(["1차 검사", "2차 검사", "3차 검사", "K 1차 검사", "Sample", "완불재검"], "category", 6)

            render_nav_buttons(step, 3)

        elif step == 3:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 도장 정보</h4><br>", unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                with c1: 
                    st.markdown("**도장일**")
                    p_date_val = st.session_state.get("painting_date", datetime.now(timezone(timedelta(hours=9))).date())
                    st.session_state.painting_date = st.date_input("도장일", value=p_date_val, label_visibility="collapsed")
                with c2: 
                    st.markdown("**도장라인**")
                    p_line_val = st.session_state.get("painting_line", "A Line")
                    st.session_state.painting_line = st.selectbox("도장라인", ["A Line", "B Line", "C Line"], index=["A Line", "B Line", "C Line"].index(p_line_val) if p_line_val in ["A Line", "B Line", "C Line"] else 0, label_visibility="collapsed")
                with c3: 
                    st.markdown("**도장순서**")
                    p_order_val = st.session_state.get("painting_order", "")
                    display_val = str(p_order_val) if p_order_val != "" else "입력"
                    if st.button(display_val, key="btn_paint_order", use_container_width=True):
                        st.session_state.numpad_buffer = ""
                        numpad_dialog("painting_order", "도장순서")
                with c4: st.write("")
        
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 조립 정보</h4><br>", unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                with c1: 
                    st.write("")
                with c2: 
                    st.markdown("**CLIP**")
                    c2_1, c2_2 = st.columns([0.45, 0.55])
                    with c2_1:
                        c_val = st.session_state.get("clip_val", "1")
                        if st.button(str(c_val) if c_val != "" else "입력", key="btn_clip", use_container_width=True):
                            st.session_state.numpad_buffer = ""
                            numpad_dialog("clip_val", "CLIP")
                    with c2_2:
                        st.checkbox("K", key="clip_k")
                with c3: 
                    st.markdown("**BASE**")
                    c3_1, c3_2 = st.columns([0.45, 0.55])
                    with c3_1:
                        b_val = st.session_state.get("base_val", "1")
                        if st.button(str(b_val) if b_val != "" else "입력", key="btn_base", use_container_width=True):
                            st.session_state.numpad_buffer = ""
                            numpad_dialog("base_val", "BASE")
                    with c3_2:
                        st.checkbox("K", key="base_k")
                with c4: 
                    st.markdown("**COVER**")
                    c4_1, c4_2 = st.columns([0.45, 0.55])
                    with c4_1:
                        cv_val = st.session_state.get("cover_val", "1")
                        if st.button(str(cv_val) if cv_val != "" else "입력", key="btn_cover", use_container_width=True):
                            st.session_state.numpad_buffer = ""
                            numpad_dialog("cover_val", "COVER")
                    with c4_2:
                        st.checkbox("K", key="cover_k")

            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 조립기 정보</h4><br>", unsafe_allow_html=True)
                render_grid_buttons(["1호기", "2호기", "3호기", "4호기", "5호기", "6호기"], "assembler_val", 6)

            st.markdown("<br>", unsafe_allow_html=True)
            c_nav = st.columns(6)
            with c_nav[4]:
                if st.button("⬅️ 이전", use_container_width=True):
                    st.session_state.step -= 1
                    st.rerun()
            with c_nav[5]:
                if st.button("작업시작 등록", type="primary", use_container_width=True):
                    uid = st.session_state.get("unique_id", "")
                    if not uid:
                        st.warning("1단계에서 바코드를 스캔하여 '고유 ID'를 생성해주세요.")
                    else:
                        with st.spinner("작업 시작 정보 등록 중..."):
                            start_dt = datetime.combine(st.session_state.get("start_date"), st.session_state.get("start_time"))
                            w_date = st.session_state.get("work_date")
                            fmt_date = f"{w_date.month}/{w_date.day}"
                            p_date = st.session_state.get("painting_date")
                            fmt_paint_date = f"{p_date.month}/{p_date.day}" 
                            in_date = st.session_state.get("in_date_field")
                            fmt_in_date = in_date.strftime("%Y-%m-%d") 
                            
                            p_line = st.session_state.get("painting_line", "")
                            fmt_paint_line = p_line.replace(" Line", "") if p_line != "선택안함" else ""
                            a_val = st.session_state.get("assembler_val", "")
                            fmt_assembler = a_val.replace("호기", "") if a_val != "선택안함" else ""
                            fmt_worker = st.session_state.get("worker", "")
                            
                            c_val = st.session_state.get("clip_val", "")
                            b_val = st.session_state.get("base_val", "")
                            cv_val = st.session_state.get("cover_val", "")
                            fmt_clip = f"K{c_val}" if st.session_state.get("clip_k") and c_val != "" else str(c_val)
                            fmt_base = f"K{b_val}" if st.session_state.get("base_k") and b_val != "" else str(b_val)
                            fmt_cover = f"K{cv_val}" if st.session_state.get("cover_k") and cv_val != "" else str(cv_val)
                            
                            lot_in = st.session_state.get("lot_input_field", "")
                            fmt_lot = f"'{lot_in}" if lot_in else ""

                            new_data = pd.DataFrame([{
                                "고유 ID": uid, "상태": "진행중",
                                "날짜": fmt_date, "교대": st.session_state.get("shift_type", ""),
                                "시작시간": start_dt.strftime("%H:%M"), "종료시간": "", "휴동시간": "", "소요시간": "", 
                                "구분": st.session_state.get("category", ""), "호기": st.session_state.get("unit", ""), 
                                "모델명(MI)": st.session_state.get("model_name", ""), "도금구분": st.session_state.get("plating_type", ""), 
                                "UPH": "", "UPD": "", "검사 수량": "", "양품수량": "", "양품 수량(전/배 포함)": "", "불량수량": "",
                                "양품율": "", "양품율(전/배 포함)": "", "완전불량율": "", "전면불량율": "", "배면불량율": "",
                                "완전불량": "", "전면불량": "", "배면불량": "", "옵셋불량": "", "수량부족": "", "기타": "", "OQC": "", "비고": "", 
                                "도장라인": fmt_paint_line, "도장일": fmt_paint_date, "도장순서": st.session_state.get("painting_order", ""), 
                                "입고일": fmt_in_date, "LOT NO.": fmt_lot, 
                                "CLIP": fmt_clip, "BASE": fmt_base, "COVER": fmt_cover, 
                                "조립기": fmt_assembler, "월": f"{w_date.month}월", "작업자": fmt_worker
                            }])
                            
                            if save_data_append(new_data):
                                st.markdown("<div style='background-color: #FFC000; color: #000000; padding: 20px; border-radius: 10px; text-align: center; font-size: 1.5rem; font-weight: 900; box-shadow: 0 4px 10px rgba(0,0,0,0.2); margin-bottom: 20px;'>✅ 새로운 작업이 진행중 상태로 등록되었습니다!</div>", unsafe_allow_html=True)
                                
                                for k, v in default_state.items():
                                    if k not in ["app_mode", "current_page", "step", "unlocked"]:
                                        st.session_state[k] = v

                                time.sleep(1.5) 
                                st.cache_data.clear() 
                                st.session_state.app_mode = "END"
                                st.session_state.step = 1
                                st.rerun()

    # ==========================================
    # 💡 2. [작업 마감] 모드 
    # ==========================================
    elif st.session_state.app_mode == "END":
        df_all = load_data()
        in_progress_df = df_all[df_all['상태'] == '진행중'].copy()

        if step == 1:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 대상 LOT 선택</h4><br>", unsafe_allow_html=True)
                if in_progress_df.empty:
                    st.info("현재 대기 중인 작업(진행중 Lot)이 없습니다.")
                    target_row = None
                else:
                    sel_col1, sel_col2 = st.columns(2)
                    with sel_col1:
                        models_in_progress = in_progress_df['모델명(MI)'].unique().tolist()
                        selected_model = st.selectbox("■ 모델명 선택", models_in_progress)
                    
                    filtered_lots = in_progress_df[in_progress_df['모델명(MI)'] == selected_model]
                    options = filtered_lots['고유 ID'].tolist()
                    def format_option(uid):
                        row = filtered_lots[filtered_lots['고유 ID'] == uid].iloc[0]
                        return f"LOT: {row['LOT NO.']} (시작: {row['시작시간']})"
                    
                    with sel_col2:
                        selected_id = st.selectbox("■ 마감할 LOT 선택", options, format_func=format_option)
                        st.session_state.target_unique_id = selected_id
                    
                    target_row = filtered_lots[filtered_lots['고유 ID'] == selected_id].iloc[0]
                    st.markdown(f"<div style='background-color: #FFC000; color: #000000; padding: 20px; border-radius: 10px; font-size: 1.2rem; font-weight: bold; box-shadow: 0 4px 10px rgba(0,0,0,0.1); margin-top: 15px;'>📌 모델명: {target_row['모델명(MI)']} &nbsp;|&nbsp; LOT: {target_row['LOT NO.']} &nbsp;|&nbsp; 시작시간: {target_row['시작시간']}</div>", unsafe_allow_html=True)

            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 작업 종료</h4><br>", unsafe_allow_html=True)
                c1, c2, c3, c4 = st.columns(4)
                with c1: 
                    st.markdown("**종료일**")
                    e_date_val = st.session_state.get("end_date", datetime.now(timezone(timedelta(hours=9))).date())
                    st.session_state.end_date = st.date_input("종료일", value=e_date_val, label_visibility="collapsed")
                with c2: 
                    st.markdown("**종료시간**")
                    end_time_obj = st.session_state.get("end_time")
                    display_val = end_time_obj.strftime("%H:%M") if end_time_obj else "입력"
                    if st.button(display_val, key="btn_end_time", use_container_width=True):
                        st.session_state.timepad_buffer = ""
                        timepad_dialog("end_time", "종료시간")
                with c3: 
                    st.markdown("**휴동시간 (분)**")
                    i_time = st.session_state.get("idle_time", 0)
                    if st.button(f"{i_time:,}", key="btn_idle_time", use_container_width=True):
                        st.session_state.numpad_buffer = str(i_time) if i_time != 0 else ""
                        numpad_dialog("idle_time", "휴동시간 (분)")
                with c4:
                    if target_row is not None:
                        try:
                            m, d = map(int, target_row['날짜'].split('/'))
                            y = st.session_state.get("work_date", datetime.now()).year
                            s_date = datetime(y, m, d).date()
                            s_time = datetime.strptime(target_row['시작시간'], "%H:%M").time()
                            start_dt = datetime.combine(s_date, s_time)
                            end_dt = datetime.combine(st.session_state.get("end_date"), st.session_state.get("end_time"))
                            if end_dt < start_dt: end_dt += timedelta(days=1)
                            
                            raw_duration = int((end_dt - start_dt).total_seconds() / 60)
                            duration_minutes = max(0, raw_duration - st.session_state.get("idle_time", 0))
                        except: duration_minutes = 0
                    else:
                        duration_minutes = 0
                    
                    st.markdown("**소요시간 (차감됨)**")
                    st.text_input("소요시간", value=f"{duration_minutes:,} 분", disabled=True, label_visibility="collapsed")

            st.markdown("<br>", unsafe_allow_html=True)
            c_nav = st.columns(6)
            with c_nav[0]:
                if st.button("🔄 신규 작업 등록", use_container_width=True):
                    st.session_state.app_mode = "START"
                    st.session_state.step = 1
                    st.rerun()
            with c_nav[5]:
                if st.button("다음 ➡️", use_container_width=True):
                    st.session_state.step = 2
                    st.rerun()

        elif step == 2:
            bad_qty = st.session_state.get("comp_def", 0) + st.session_state.get("front_def", 0) + st.session_state.get("rear_def", 0) + st.session_state.get("offset_def", 0) + st.session_state.get("etc_def", 0)
            total_qty = max(0, st.session_state.get("good_qty", 0) + bad_qty - st.session_state.get("shortage_qty", 0))

            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 수량 등록</h4><br>", unsafe_allow_html=True)
                q1, q2, q3, q4 = st.columns(4)
                with q1: 
                    st.markdown("**검사 수량 (자동)**")
                    st.text_input("검사 수량", value=f"{total_qty:,}", disabled=True, label_visibility="collapsed")
                with q2: 
                    st.markdown("**양품수량**")
                    g_qty = st.session_state.get("good_qty", 0)
                    if st.button(f"{g_qty:,}", key="f_good", use_container_width=True): 
                        val = str(g_qty)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("good_qty", "양품수량")
                with q3: 
                    st.markdown("**불량수량 (자동)**")
                    st.text_input("불량수량", value=f"{bad_qty:,}", disabled=True, label_visibility="collapsed")
                with q4:
                    st.markdown("**OQC**")
                    o_val = st.session_state.get("oqc_status", "선택안함")
                    st.session_state.oqc_status = st.selectbox("OQC", ["선택안함", "육안", "OQC"], index=["선택안함", "육안", "OQC"].index(o_val) if o_val in ["선택안함", "육안", "OQC"] else 0, label_visibility="collapsed")
            
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 불량 세부 내역</h4><br>", unsafe_allow_html=True)
                c1, c2, c3, c4, c5, c6 = st.columns(6)
                with c1: 
                    st.markdown("**완전불량**")
                    c_def = st.session_state.get("comp_def", 0)
                    if st.button(f"{c_def:,}", key="f_comp", use_container_width=True): 
                        val = str(c_def)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("comp_def", "완전불량")
                with c2: 
                    st.markdown("**전면불량**")
                    f_def = st.session_state.get("front_def", 0)
                    if st.button(f"{f_def:,}", key="f_front", use_container_width=True): 
                        val = str(f_def)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("front_def", "전면불량")
                with c3: 
                    st.markdown("**배면불량**")
                    r_def = st.session_state.get("rear_def", 0)
                    if st.button(f"{r_def:,}", key="f_rear", use_container_width=True): 
                        val = str(r_def)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("rear_def", "배면불량")
                with c4: 
                    st.markdown("**옵셋불량**")
                    o_def = st.session_state.get("offset_def", 0)
                    if st.button(f"{o_def:,}", key="f_off", use_container_width=True): 
                        val = str(o_def)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("offset_def", "옵셋불량")
                with c5: 
                    st.markdown("**수량부족**")
                    s_qty = st.session_state.get("shortage_qty", 0)
                    if st.button(f"{s_qty:,}", key="f_short", use_container_width=True): 
                        val = str(s_qty)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("shortage_qty", "수량부족")
                with c6: 
                    st.markdown("**기타**")
                    e_def = st.session_state.get("etc_def", 0)
                    if st.button(f"{e_def:,}", key="f_etc", use_container_width=True): 
                        val = str(e_def)
                        st.session_state.numpad_buffer = val if val != "0" else ""
                        numpad_dialog("etc_def", "기타")

            if total_qty > 0:
                comp_rate = (st.session_state.get("comp_def", 0) / total_qty) * 100
                front_rate = (st.session_state.get("front_def", 0) / total_qty) * 100
                rear_rate = (st.session_state.get("rear_def", 0) / total_qty) * 100
                offset_rate = (st.session_state.get("offset_def", 0) / total_qty) * 100
                
                if comp_rate > 5.0 and not st.session_state.get("comp_warned", False):
                    show_sbl_warning("완전불량", comp_rate)
                    st.session_state.comp_warned = True
                if front_rate > 5.0 and not st.session_state.get("front_warned", False):
                    show_sbl_warning("전면불량", front_rate)
                    st.session_state.front_warned = True
                if rear_rate > 5.0 and not st.session_state.get("rear_warned", False):
                    show_sbl_warning("배면불량", rear_rate)
                    st.session_state.rear_warned = True
                if offset_rate > 5.0 and not st.session_state.get("offset_warned", False):
                    show_sbl_warning("옵셋불량", offset_rate)
                    st.session_state.offset_warned = True

            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 실시간 수율 현황</h4><br>", unsafe_allow_html=True)
                rate_good = round((st.session_state.get("good_qty", 0) / total_qty) * 100, 1) if total_qty > 0 else 0.0
                
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
                        round((st.session_state.get("comp_def", 0)/total_qty)*100,1) if total_qty>0 else 0,
                        round((st.session_state.get("front_def", 0)/total_qty)*100,1) if total_qty>0 else 0,
                        round((st.session_state.get("rear_def", 0)/total_qty)*100,1) if total_qty>0 else 0,
                        round((st.session_state.get("offset_def", 0)/total_qty)*100,1) if total_qty>0 else 0
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
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 최종 확인 및 저장</h4><br>", unsafe_allow_html=True)
                rem_col, nav_col1, nav_col2 = st.columns([0.6, 0.2, 0.2])
                with rem_col:
                    st.markdown("**비고**")
                    st.session_state.remarks = st.text_area("비고", value=st.session_state.get("remarks", ""), label_visibility="collapsed")
                    
                with nav_col1:
                    st.markdown("**&nbsp;**") 
                    if st.button("⬅️ 이전", use_container_width=True):
                        st.session_state.step -= 1
                        st.rerun()
                with nav_col2:
                    st.markdown("**&nbsp;**") 
                    if st.button("Data 최종 저장", type="primary", use_container_width=True):
                        if total_qty == 0: st.warning("입력된 수량 데이터가 없습니다.")
                        elif not st.session_state.get("target_unique_id", ""): st.warning("1단계에서 마감할 Lot를 선택해주세요.")
                        else:
                            with st.spinner("DB 마감 업데이트 중..."):
                                df_all = load_data()
                                target_id = st.session_state.target_unique_id
                                df_target = df_all[df_all['고유 ID'] == target_id]
                                
                                if df_target.empty:
                                    st.error("오류: DB에서 해당 Lot를 찾을 수 없습니다.")
                                else:
                                    target_idx = df_target.index[0]
                                    target_row = df_target.iloc[0].to_dict()
                                    
                                    try:
                                        m, d = map(int, target_row['날짜'].split('/'))
                                        y = st.session_state.get("work_date", datetime.now()).year
                                        s_date = datetime(y, m, d).date()
                                        s_time = datetime.strptime(target_row['시작시간'], "%H:%M").time()
                                        start_dt = datetime.combine(s_date, s_time)
                                        end_dt = datetime.combine(st.session_state.get("end_date"), st.session_state.get("end_time"))
                                        if end_dt < start_dt: end_dt += timedelta(days=1)
                                        
                                        raw_duration = int((end_dt - start_dt).total_seconds() / 60)
                                        duration_minutes = max(0, raw_duration - st.session_state.get("idle_time", 0))
                                    except:
                                        duration_minutes = 0

                                    uph_val = int((total_qty / duration_minutes) * 60) if duration_minutes > 0 else 0
                                    upd_val = uph_val * 22
                                    good_include_front_rear = st.session_state.get("good_qty", 0) + st.session_state.get("front_def", 0) + st.session_state.get("rear_def", 0)
                                    
                                    if total_qty > 0:
                                        rate_good = round((st.session_state.get("good_qty", 0) / total_qty) * 100, 1)
                                        rate_good_inc = round((good_include_front_rear / total_qty) * 100, 1)
                                        comp_rate_num = round(st.session_state.get("comp_def", 0) / total_qty * 100, 1)
                                        front_rate_num = round(st.session_state.get("front_def", 0) / total_qty * 100, 1)
                                        rear_rate_num = round(st.session_state.get("rear_def", 0) / total_qty * 100, 1)
                                        offset_rate_num = round(st.session_state.get("offset_def", 0) / total_qty * 100, 1)
                                    else:
                                        rate_good = rate_good_inc = comp_rate_num = front_rate_num = rear_rate_num = offset_rate_num = 0.0

                                    target_row.update({
                                        "상태": "완료",
                                        "종료시간": st.session_state.get("end_time").strftime("%H:%M"),
                                        "휴동시간": f"{st.session_state.get('idle_time', 0):,}", 
                                        "소요시간": f"{duration_minutes:,}",
                                        "UPH": f"{uph_val:,}", "UPD": f"{upd_val:,}",
                                        "검사 수량": f"{total_qty:,}", "양품수량": f"{st.session_state.get('good_qty', 0):,}", "양품 수량(전/배 포함)": f"{good_include_front_rear:,}", "불량수량": f"{bad_qty:,}",
                                        "양품율": f"{rate_good:.1f}%", "양품율(전/배 포함)": f"{rate_good_inc:.1f}%",
                                        "완전불량율": f"{comp_rate_num:.1f}%", "전면불량율": f"{front_rate_num:.1f}%", "배면불량율": f"{rear_rate_num:.1f}%",
                                        "완전불량": f"{st.session_state.get('comp_def', 0):,}", "전면불량": f"{st.session_state.get('front_def', 0):,}", "배면불량": f"{st.session_state.get('rear_def', 0):,}", "옵셋불량": f"{st.session_state.get('offset_def', 0):,}", "수량부족": f"{st.session_state.get('shortage_qty', 0):,}", "기타": f"{st.session_state.get('etc_def', 0):,}",
                                        "OQC": "" if st.session_state.get("oqc_status", "선택안함") == "선택안함" else st.session_state.get("oqc_status"), 
                                        "비고": st.session_state.get("remarks", "")
                                    })
                                    
                                    sheet = get_sheet()
                                    if sheet:
                                        sheet_row = target_row['_sheet_row'] 
                                        updated_row_list = ["nan" if pd.isna(target_row.get(col, "")) else "" if str(target_row.get(col, "")).strip().lower() in ["nan", "none"] else str(target_row.get(col, "")).strip() for col in EXCEL_COLUMNS]
                                        sheet.update(values=[updated_row_list], range_name=f'A{sheet_row}')
                                        
                                        st.markdown("<div style='background-color: #FFC000; color: #000000; padding: 20px; border-radius: 10px; text-align: center; font-size: 1.5rem; font-weight: 900; box-shadow: 0 4px 10px rgba(0,0,0,0.2); margin-bottom: 20px;'>✅ 데이터가 성공적으로 마감되었습니다!</div>", unsafe_allow_html=True)
                                        st.cache_data.clear()
                                        
                                        for k, v in default_state.items():
                                            if k not in ["app_mode", "current_page", "step", "unlocked"]:
                                                st.session_state[k] = v
                                                
                                        time.sleep(1.5)
                                        st.session_state.app_mode = "EDIT"
                                        st.session_state.step = 1
                                        st.rerun()

    # ==========================================
    # 💡 3. [데이터 수정] 모드
    # ==========================================
    elif st.session_state.app_mode == "EDIT":
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 최근 저장 Data List</h4><br>", unsafe_allow_html=True)
            df_history = load_data().copy()
            
            if not df_history.empty:
                df_history['orig_index'] = df_history['_sheet_row']
                recent_20 = df_history.iloc[::-1].head(20).copy()
                display_df = recent_20.drop(columns=['orig_index', '_sheet_row'])
                
                edited_df = st.data_editor(
                    display_df, 
                    use_container_width=True, 
                    hide_index=True,
                    column_config={"LOT NO.": st.column_config.TextColumn("LOT NO.")}
                )
                
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("Data 수정 적용", type="primary", use_container_width=True):
                    sheet = get_sheet()
                    changed = False
                    with st.spinner("구글 시트 업데이트 중..."):
                        for idx in edited_df.index:
                            old_row = display_df.loc[idx].fillna("").astype(str).tolist()
                            new_row = edited_df.loc[idx].fillna("").astype(str).tolist()
                            if old_row != new_row:
                                gspread_row = recent_20.loc[idx, 'orig_index']
                                sheet.update(values=[new_row], range_name=f'A{gspread_row}')
                                changed = True
                    
                    if changed:
                        st.markdown("<div style='background-color: #FFC000; color: #000000; padding: 20px; border-radius: 10px; text-align: center; font-size: 1.5rem; font-weight: 900; box-shadow: 0 4px 10px rgba(0,0,0,0.2); margin-bottom: 20px;'>✅ 구글 시트에 수정 내용이 성공적으로 반영되었습니다!</div>", unsafe_allow_html=True)
                        st.cache_data.clear()
                        time.sleep(1.5)
                        st.rerun()
                    else:
                        st.info("수정된 항목이 없습니다.")
            else:
                st.caption("저장된 데이터가 없습니다.")
