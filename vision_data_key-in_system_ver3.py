import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import json
import os
import re
import base64
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="VISION DATA KEY-IN SYSTEM", layout="wide", initial_sidebar_state="collapsed")

# 💡 이미지 로드 헬퍼 함수
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

# 💡 시스템 상태 초기화
if "main_authenticated" not in st.session_state: st.session_state.main_authenticated = False

# 💡 프리미엄 UI 및 [메뉴/배지 숨김 처리 CSS]
global_theme_css = """
<style>
/* 🚫 Streamlit 기본 상단 헤더, 메뉴, 툴바 완벽 은닉 */
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { display: none !important; visibility: hidden !important; }
[data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
footer { display: none !important; } 

/* 🚫 사이드바 및 붕 뜨는 공간 제거 */
[data-testid="collapsedControl"] { display: none !important; pointer-events: none !important; }
[data-testid="stSidebar"] { display: none !important; }
body { overscroll-behavior-y: none !important; background-color: #f8fafc !important; } 
::-webkit-scrollbar { display: none; }
.block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 98% !important; }

/* 💡 강제 라이트 테마 */
h1, h2, h3, h4, h5, h6, p, label { font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif !important; color: #1e293b !important; }
[data-testid="stAppViewContainer"] { background-color: #f8fafc !important; color: #1e293b !important; }
div[data-baseweb="input"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }
div[data-baseweb="input"] input { color: #1e293b !important; font-weight: bold !important; }

div[data-testid="stVerticalBlockBorderWrapper"] { background-color: #ffffff !important; border-radius: 12px !important; border: 1px solid #e2e8f0 !important; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.03) !important; padding: 1.5rem !important; margin-bottom: 0.8rem !important; }
.command-header { color: #1e293b !important; font-weight: 900 !important; letter-spacing: 1px; }

div[data-testid="stButton"] button { height: 2.6rem !important; min-height: 2.6rem !important; font-size: 1.1rem !important; font-weight: bold !important; border-radius: 8px !important; background-color: #E7E6E6 !important; color: #000000 !important; border: 1px solid #cbd5e1 !important; transition: all 0.2s ease; }
div[data-testid="stButton"] button p { color: #000000 !important; }
div[data-testid="stButton"] button:hover { background-color: #1e293b !important; border-color: #1e293b !important; }
div[data-testid="stButton"] button:hover p { color: #ffffff !important; }
div[data-testid="stButton"] button[kind="primary"] { background-color: #1e293b !important; border: 1px solid #0f172a !important; }
div[data-testid="stButton"] button[kind="primary"]:hover { background-color: #0f172a !important; }
div[data-testid="stButton"] button[kind="primary"] p { color: #ffffff !important; }
</style>
"""
st.markdown(global_theme_css, unsafe_allow_html=True)

# 🛡️ 전역 방어막: 로그인 화면을 포함한 "모든 페이지"에서 Manage App 및 프로필 뱃지 원천 차단
badge_killer_script = """
<script>
const nukeManageApp = () => {
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
nukeManageApp();
setInterval(nukeManageApp, 100); 
</script>
"""
components.html(badge_killer_script, height=0, width=0)

# ==========================================
# 💡 메인 시스템 전용 로그인 페이지 (logo.png 적용)
# ==========================================
if not st.session_state.main_authenticated:
    st.markdown("<br><br><br><br><br>", unsafe_allow_html=True)
    col_sp1, col_auth, col_sp3 = st.columns([1, 1, 1])
    with col_auth:
        with st.container(border=True):
            logo_l_data = get_image_base64("logo")
            if logo_l_data:
                st.markdown(f"<div style='text-align: center;'><img src='{logo_l_data}' style='max-width: 100%; max-height: 80px; object-fit: contain; margin-bottom: 15px;'></div>", unsafe_allow_html=True)
                
            st.markdown("<h3 style='text-align:center; color:#1e293b; font-weight:900;'>🔐 KEY-IN SYSTEM 인증</h3>", unsafe_allow_html=True)
            st.markdown("<div style='text-align:center; color:#64748b; margin-bottom:20px; font-weight:bold;'>데이터 입력 권한을 위해 비밀번호를 입력하세요.</div>", unsafe_allow_html=True)
            pwd = st.text_input("비밀번호", type="password", label_visibility="collapsed", placeholder="비밀번호 입력", key="main_pwd")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("✅ 접속", type="primary", use_container_width=True, key="main_confirm"):
                if pwd == "7777":  # 필요 시 메인 시스템 비밀번호로 변경 가능
                    st.session_state.main_authenticated = True
                    st.rerun()
                else:
                    st.error("비밀번호가 일치하지 않습니다.")
    st.stop()

# ==========================================
# 💡 메인 데이터 입력 및 관리 시스템 본문 영역
# ==========================================
st.markdown("<div class='command-header' style='font-size: 1.8rem; margin-bottom: 20px;'>📝 VISION DATA KEY-IN WIZARD</div>", unsafe_allow_html=True)

# 여기에 기존에 작성하셨던 데이터 입력 폼, 구글 시트 연동, 설정 관리(VIEWER_CONFIG 저장 등) 코드를 이어 붙여 사용하시면 됩니다.
st.success("🎉 인증이 완료되었습니다. 안전하게 데이터를 입력하고 관리하세요!")
