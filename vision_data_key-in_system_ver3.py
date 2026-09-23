import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import json
import os
import re
import base64
import time
import uuid
from datetime import datetime, time as dt_time, timedelta, timezone
from io import BytesIO
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

worker_list = ["작업자 선택", "한상일", "지한구", "노준혁", "이명희", "조난희", "김영민", "송민재", "배현정", "김환용", "허건", "김현정", "관리자"]
model_list = ["D65S(KRIOS)", "MEM", "Centaur", "Sphinx-E", "Banff", "AV-J", "Seattle", "Juliet-O"]

st.set_page_config(page_title="VISION DATA KEY-IN SYSTEM", layout="wide", initial_sidebar_state="expanded")

# ==========================================
# 💡 이미지 & 헬퍼 함수 모음
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
            if not os.path.exists(directory): continue
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

# ==========================================
# 💡 구글 시트 연동 및 데이터 관리 함수
# ==========================================
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
        try: return doc.worksheet(TAB_NAME)
        except Exception as e: return doc.sheet1
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
        return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row', 'DateTime', 'DateOnly'])
    
    if len(raw_data) < 24: return pd.DataFrame(columns=EXCEL_COLUMNS + ['_sheet_row', 'DateTime', 'DateOnly'])
    
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
    df['_sheet_row'] = range(data_start_idx + 1, data_start_idx + 1 + len(df))
    
    rename_dict = {}
    mapped_std_cols = set()
    for c_idx, h in enumerate(unique_headers):
        cc = h.replace(" ", "").replace("률", "율").upper()
        matched_col = None
        if c_idx == 2: matched_col = '날짜'
        elif "고유" in cc and "ID" in cc: matched_col = '고유 ID'
        elif cc == "상태": matched_col = '상태'
        elif cc == "교대": matched_col = '교대'
        elif cc == "시작시간": matched_col = '시작시간'
        elif cc == "종료시간": matched_col = '종료시간'
        elif cc == "휴동시간": matched_col = '휴동시간'
        elif cc == "소요시간": matched_col = '소요시간'
        elif cc == "구분": matched_col = '구분'
        elif cc == "호기": matched_col = '호기'
        elif cc in ["모델명(MI)", "모델명", "품명", "MI"]: matched_col = '모델명(MI)'
        elif cc in ["도금구분", "도금"]: matched_col = '도금구분'
        elif cc == "UPH": matched_col = 'UPH'
        elif cc == "UPD": matched_col = 'UPD'
        elif cc in ["검사수량", "총수량"]: matched_col = '검사 수량'
        elif cc == "양품수량": matched_col = '양품수량'
        elif "양품수량" in cc and "포함" in cc: matched_col = '양품 수량(전/배 포함)'
        elif cc == "불량수량": matched_col = '불량수량'
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
        elif cc == "수량부족": matched_col = '수량부족'
        elif cc == "기타": matched_col = '기타'
        elif cc in ["육안/OQC", "OQC"]: matched_col = 'OQC'
        elif cc == "비고": matched_col = '비고'
        elif cc == "도장라인": matched_col = '도장라인'
        elif cc == "도장일": matched_col = '도장일'
        elif cc == "도장순서": matched_col = '도장순서'
        elif cc == "입고일": matched_col = '입고일'
        elif cc in ["LOTNO.", "LOTNO", "LOT", "로트"]: matched_col = 'LOT NO.'
        elif cc == "CLIP": matched_col = 'CLIP'
        elif cc == "BASE": matched_col = 'BASE'
        elif cc == "COVER": matched_col = 'COVER'
        elif cc == "조립기": matched_col = '조립기'
        elif cc == "월": matched_col = '월'
        elif cc == "작업자": matched_col = '작업자'
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
        
    parsed_dates = df.apply(parse_dt, axis=1)
    missing_dates_idx = parsed_dates.isna()
    if missing_dates_idx.any():
        parsed_dates.loc[missing_dates_idx] = [datetime(2026, 1, 1) + timedelta(minutes=i) for i in range(missing_dates_idx.sum())]
    df['DateTime'] = pd.to_datetime(parsed_dates)
    df['DateOnly'] = df['DateTime'].dt.date 

    if '구분' in df.columns:
        df_filtered = df[df['구분'].fillna('').astype(str).str.contains('1차', na=False)]
        if not df_filtered.empty: df = df_filtered
        
    return df[ext_cols + ['_sheet_row', 'DateTime', 'DateOnly']]

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

# ==========================================
# 💡 키인 폼 입력용 다이얼로그 및 콜백 함수
# ==========================================
def parse_scanned_data():
    raw_val = st.session_state.get("scanned_raw_data", "")
    if not raw_val: return
    st.session_state.unique_id = raw_val
    if '$' in raw_val:
        parts = [p for p in raw_val.split('$') if p]
        if len(parts) >= 5: st.session_state.lot_input_field = parts[4]
        else: st.session_state.lot_input_field = parts[-1]
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

@st.dialog("🚨 SBL 관리 한계 초과 알림")
def show_sbl_warning(defect_name, rate, limit_val):
    st.markdown(f"""
    <div style='text-align:center; padding: 20px; background-color: #fef2f2; border: 3px solid #ef4444; border-radius: 12px;'>
        <h3 style='color: #b91c1c; margin-top: 0; font-weight: 900;'>[{defect_name}] SBL 기준치 초과!</h3>
        <span style='font-size: 3.5rem; font-weight: 900; color: #ef4444;'>{rate:.1f}%</span><br><br>
        <span style='color: #1e293b; font-size: 1.1rem; font-weight: bold;'>설정된 관리 기준({limit_val:.1f}%)을 초과했습니다.<br>즉시 관리자에게 보고하고 해당 LOT를 확인하세요.</span>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("확인 및 인지 완료 (닫기)", type="primary", use_container_width=True):
        st.rerun()

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

# ==========================================
# 💡 세션 상태 및 설정 동기화
# ==========================================
if "unlocked" not in st.session_state: st.session_state.unlocked = False
if "sys_menu" not in st.session_state: st.session_state.sys_menu = "keyin"
if "app_mode" not in st.session_state: st.session_state.app_mode = "START" 
if "step" not in st.session_state: st.session_state.step = 1
if "rotate_idx" not in st.session_state: st.session_state.rotate_idx = 0
if "edit_unlocked" not in st.session_state: st.session_state.edit_unlocked = False
if "unlocked" in st.query_params:
    st.session_state.unlocked = True
    st.query_params.clear()

config = load_shared_config() or {}
st.session_state.sbl_limits = config.get("sbl_limits", {
    "Yield_Default": 85.0, "Yield_Centaur": 91.4, "Yield_MEM": 93.2,
    "Def_Comp": 10.0, "Def_Front": 5.0, "Def_Rear": 5.0, "Def_Offset": 5.0
})
if 'sel_std' not in st.session_state: st.session_state.sel_std = config.get("sel_std", [])
if 'sel_inc' not in st.session_state: st.session_state.sel_inc = config.get("sel_inc", [])

default_state = {
    "unique_id": "", "work_date": datetime.now(timezone(timedelta(hours=9))).date(), 
    "shift_type": "주간", "worker": None,
    "model_name": "D65S(KRIOS)", "lot_input_field": "", "in_date_field": datetime.now(timezone(timedelta(hours=9))).date(),
    "plating_type": "A", "start_date": datetime.now(timezone(timedelta(hours=9))).date(), "start_time": datetime.now(timezone(timedelta(hours=9))).time(),
    "end_date": datetime.now(timezone(timedelta(hours=9))).date(), "end_time": datetime.now(timezone(timedelta(hours=9))).time(), "unit": "1호기",
    "category": "1차 검사", "idle_time": 0, "painting_date": datetime.now(timezone(timedelta(hours=9))).date(),
    "painting_order": "", "painting_line": "A Line", 
    "clip_val": "1", "clip_type": "일반",
    "base_val": "1", "cover_val": "1",
    "assembler_val": "1호기",
    "good_qty": 0, "comp_def": 0, "front_def": 0, "rear_def": 0, "offset_def": 0,
    "shortage_qty": 0, "etc_def": 0, "oqc_status": "선택안함", "remarks": "",
    "scanned_raw_data": "", "comp_warned": False, "front_warned": False, 
    "rear_warned": False, "offset_warned": False,
    "numpad_buffer": "", "timepad_buffer": "", "target_unique_id": ""
}
for key, value in default_state.items():
    if key not in st.session_state: st.session_state[key] = value

# ==========================================
# 🛡️ 전역 CSS (깜빡임 완벽 차단 & UI 최적화)
# ==========================================
global_theme_css = """
<style>
/* 🚫 헤더 및 상단 메뉴, 툴바 완벽 은닉 */
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { display: none !important; visibility: hidden !important; }
[data-testid="stToolbar"] { display: none !important; visibility: hidden !important; }
footer { display: none !important; } 

/* 🚫 Streamlit Deploy, 뱃지 등 강제 은닉 (1차 CSS 방어) */
[data-testid="manage-app-button"],
[data-testid="stAppDeployButton"],
.stDeployButton,
div[class^="viewerBadge"],
div[class*="viewerBadge"],
#creatorBadge {
    display: none !important; opacity: 0 !important; visibility: hidden !important;
    z-index: -1000 !important; pointer-events: none !important;
}

/* iframe 텍스트 깜빡임 방지 (크기 0으로 축소) */
iframe[title="streamlit_components.components.html"] {
    display: none !important; opacity: 0 !important; width: 0 !important; height: 0 !important; position: absolute !important;
}

/* 히든 버튼 완전 은닉 (DOM 상에서 존재를 감춤) */
.hidden-btn { display: none !important; visibility: hidden !important; height: 0px !important; width: 0px !important; overflow: hidden !important; position: absolute !important; z-index: -9999 !important; }

body { overscroll-behavior-y: none !important; background-color: #f8fafc !important; } 
::-webkit-scrollbar { display: none; }
.block-container { padding-top: 3rem !important; padding-bottom: 2rem !important; padding-left: 1.5rem !important; padding-right: 1.5rem !important; max-width: 98% !important; }

h1, h2, h3, h4, h5, h6, p, div, span, label { font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', 'Noto Sans KR', sans-serif !important; }
[data-testid="stAppViewContainer"] { background-color: #f1f5f9 !important; color: #1e293b !important; }

/* 💡 공통 버튼 스타일 */
div[data-testid="stButton"] button { height: 2.6rem !important; min-height: 2.6rem !important; font-size: 1.1rem !important; font-weight: bold !important; border-radius: 8px !important; background-color: #E7E6E6 !important; border: 1px solid #cbd5e1 !important; transition: all 0.2s ease; }
div[data-testid="stButton"] button p { color: #000000 !important; }
div[data-testid="stButton"] button:hover { background-color: #1e293b !important; border-color: #1e293b !important; }
div[data-testid="stButton"] button:hover p { color: #ffffff !important; }

div[data-testid="stButton"] button[kind="primary"] { background-color: #1e293b !important; border: 1px solid #0f172a !important; }
div[data-testid="stButton"] button[kind="primary"]:hover { background-color: #0f172a !important; }
div[data-testid="stButton"] button[kind="primary"] p { color: #ffffff !important; }

/* 🛡️ CSS 우측 하단 절대 방어막 (클릭 불가) */
.stApp::after {
    content: "" !important; position: fixed !important; bottom: 0 !important; right: 0 !important; width: 300px !important; height: 150px !important;
    background: transparent !important; z-index: 2147483647 !important; pointer-events: auto !important; cursor: default !important;
}
</style>
"""
st.markdown(global_theme_css, unsafe_allow_html=True)


# ==========================================
# 💡 잠금 화면 (슬라이더 언락)
# ==========================================
if not st.session_state.unlocked:
    st.markdown("""
    <style>
        [data-testid="stSidebar"] { display: none !important; }
        [data-testid="collapsedControl"] { display: none !important; }
    </style>
    """, unsafe_allow_html=True)
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        logo_l_data = get_image_base64("logo")
        if logo_l_data:
            # 💡 로고 최대 크기 400px 반영 및 배경 투명화
            st.markdown(f"<div style='text-align: center;'><img src='{logo_l_data}' style='max-width: 100%; max-height: 400px; object-fit: contain; margin-bottom: 20px; mix-blend-mode: multiply;'></div>", unsafe_allow_html=True)
        else:
            st.markdown("<h1 style='text-align: center; color: #1e293b; font-size: 45px; font-weight: 900; letter-spacing: 2px;'>VISION DATA KEY-IN SYSTEM</h1><br><br>", unsafe_allow_html=True)
        
        # 💡 히든 버튼 완전 격리
        hidden_container = st.empty()
        with hidden_container.container():
            st.markdown("<div class='hidden-btn'>", unsafe_allow_html=True)
            if st.button("UNLOCK_SYSTEM_BTN_HIDDEN"):
                st.session_state.unlocked = True
                st.rerun()
            st.markdown("</div>", unsafe_allow_html=True)
            
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
        # 스크립트 실행 박스를 높이 0으로 처리해 화면 밀림 방지
        components.html(slider_html, height=100, width=0)
        
    st.markdown("<div style='position: fixed; bottom: 10%; left: 0; width: 100%; text-align: center; font-size: 10pt; color: #FFC000 !important; font-weight: bold;'>Created by --- Romero.K</div>", unsafe_allow_html=True)
    st.stop()


# ==========================================
# 💡 최상단 1x2 통합 네비게이션 (8:2 비율 적용)
# ==========================================
if st.session_state.sys_menu != "exit":
    st.markdown("<h2 style='text-align: center; color: #1e293b; font-weight: 900; margin-bottom: 20px;'>VISION DATA KEY-IN SYSTEM</h2>", unsafe_allow_html=True)

    m_col1, m_col2 = st.columns([0.8, 0.2])
    with m_col1:
        if st.button("📝 KEY-IN WIZARD", use_container_width=True, type="primary" if st.session_state.sys_menu == "keyin" else "secondary"):
            st.session_state.sys_menu = "keyin"
            st.rerun()
    with m_col2:
        if st.button("⚙️ ADMINISTRATOR", use_container_width=True, type="primary" if st.session_state.sys_menu in ["admin", "viewer"] else "secondary"):
            st.session_state.sys_menu = "admin"
            st.rerun()
            
    # 💡 플로팅 EXIT 버튼용 히든 트리거 완전 격리
    hidden_exit_container = st.empty()
    with hidden_exit_container.container():
        st.markdown("<div class='hidden-btn'>", unsafe_allow_html=True)
        if st.button("HIDDEN_EXIT_TRIGGER"):
            st.session_state.sys_menu = "exit"
            st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<hr style='margin-top: 5px; margin-bottom: 20px; border-color: #cbd5e1;'>", unsafe_allow_html=True)


# ==========================================
# 🚀 [라우팅 1] KEY-IN WIZARD
# ==========================================
if st.session_state.sys_menu == "keyin":
    st.markdown("""
    <style>
    /* 사이드바 UI 100% 화이트닝 및 레이아웃 유지 */
    [data-testid="stSidebar"] { background-color: #0f172a !important; border-right: 1px solid #cbd5e1 !important; }
    [data-testid="stSidebar"] * { color: #ffffff !important; }
    [data-testid="stSidebar"] .stButton > button { height: 48px !important; max-height: 48px !important; justify-content: flex-start !important; padding-left: 15px !important; margin-bottom: 5px !important; border-radius: 6px !important; background-color: transparent !important; border: 1px solid rgba(255,255,255,0.1) !important; color: #ffffff !important; box-shadow: none !important; }
    [data-testid="stSidebar"] .stButton > button p { font-weight: 800 !important; font-size: 14px !important; text-indent: 10px !important; text-align: left !important; color: #ffffff !important; }
    [data-testid="stSidebar"] .stButton > button[kind="primary"] { background-color: #3b82f6 !important; color: #ffffff !important; border: none !important; border-left: 4px solid #FFC000 !important; }
    [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover { background-color: rgba(255,255,255,0.1) !important; color: #ffffff !important; border: 1px solid rgba(255,255,255,0.5) !important; }
    [data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover p { color: #ffffff !important; }
    
    /* 기본 숨김 해제 */
    [data-testid="collapsedControl"] { display: none !important; }
    
    /* 폼 영역 디자인 */
    div[data-testid="stVerticalBlockBorderWrapper"] { background-color: #ffffff !important; border-radius: 12px !important; border: 1px solid #cbd5e1 !important; box-shadow: 0 4px 10px rgba(0, 0, 0, 0.04) !important; padding: 1.5rem !important; margin-bottom: 0.8rem !important; }
    div[data-baseweb="input"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }
    div[data-baseweb="input"] input { color: #1e293b !important; font-weight: bold !important; }
    div[data-testid="stNumberInput"] div[data-baseweb="input"] > div { background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; border-radius: 6px; }
    div[data-testid="stNumberInput"] input { color: #1e293b !important; font-weight: bold !important; }
    </style>
    """, unsafe_allow_html=True)
    
    # 💡 사이드바 렌더링 (DATA LIST 와 관리자 기능 재배치 적용)
    with st.sidebar:
        KST = timezone(timedelta(hours=9))
        now = datetime.now(KST)
        weekdays = ['월', '화', '수', '목', '금', '토', '일']
        current_time_str = f"{now.strftime('%Y년 %m월 %d일')} ({weekdays[now.weekday()]}) {now.strftime('%p %I:%M').replace('AM', '오전').replace('PM', '오후')}"
        st.markdown(f"<div style='text-align: center; color: #ffffff !important; background-color: rgba(255,255,255,0.1) !important; padding: 10px; border-radius: 8px; font-weight: bold; font-size: 0.85rem; margin-bottom: 20px;'>{current_time_str}</div>", unsafe_allow_html=True)
        
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

        st.markdown("<br><h4 style='color: #f8fafc; font-size: 1.1rem; border-bottom: 1px solid #334155; padding-bottom: 8px;'>■ DATA LIST</h4><br>", unsafe_allow_html=True)
        if st.button("최근 저장 Data List", type="primary" if st.session_state.app_mode=="EDIT" else "secondary", use_container_width=True):
            st.session_state.app_mode = "EDIT"
            st.session_state.step = 1
            st.rerun()
            
        st.markdown("<hr style='border-color: #334155; margin-top: 10px; margin-bottom: 10px;'>", unsafe_allow_html=True)
        st.markdown("<div style='text-align: center; color: #FFC000 !important; font-size: 14px; font-weight: bold;'>Created by --- Romero.K</div>", unsafe_allow_html=True)

    step = st.session_state.step

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
                    w_val = st.session_state.get("worker", None)
                    try: w_idx = worker_list.index(w_val) if w_val in worker_list else None
                    except ValueError: w_idx = None
                    st.session_state.worker = st.selectbox("작업자", worker_list, index=w_idx, placeholder="작업자 선택", label_visibility="collapsed")

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
                    if not st.session_state.get("worker"):
                        st.warning("작업자를 선택해주세요.")
                    else:
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

            st.markdown("<br>", unsafe_allow_html=True)
            c_nav = st.columns(6)
            with c_nav[4]:
                if st.button("⬅️ 이전", use_container_width=True):
                    st.session_state.step -= 1
                    st.rerun()
            with c_nav[5]:
                if st.button("다음 ➡️", use_container_width=True):
                    st.session_state.step = 3
                    st.rerun()

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
                    clip_type = st.session_state.get("clip_type", "일반")
                    c_val = st.session_state.get("clip_val", "1")
                    if st.button(str(c_val) if c_val != "" else "입력", key="btn_clip", use_container_width=True):
                        st.session_state.numpad_buffer = ""
                        numpad_dialog("clip_val", "CLIP")
                    st.session_state.clip_type = st.radio("CLIP 옵션", ["일반", "K1", "K2", "K3"], index=["일반", "K1", "K2", "K3"].index(clip_type), horizontal=True, label_visibility="collapsed")
                with c3: 
                    st.markdown("**BASE**")
                    b_val = st.session_state.get("base_val", "1")
                    if st.button(str(b_val) if b_val != "" else "입력", key="btn_base", use_container_width=True):
                        st.session_state.numpad_buffer = ""
                        numpad_dialog("base_val", "BASE")
                with c4: 
                    st.markdown("**COVER**")
                    cv_val = st.session_state.get("cover_val", "1")
                    if st.button(str(cv_val) if cv_val != "" else "입력", key="btn_cover", use_container_width=True):
                        st.session_state.numpad_buffer = ""
                        numpad_dialog("cover_val", "COVER")

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
                            
                            clip_t = st.session_state.get("clip_type", "일반")
                            clip_v = str(st.session_state.get("clip_val", ""))
                            if clip_t == "일반":
                                fmt_clip = clip_v
                            else:
                                fmt_clip = f"{clip_t}-{clip_v}" if clip_v else clip_t
                                
                            fmt_base = str(st.session_state.get("base_val", ""))
                            fmt_cover = str(st.session_state.get("cover_val", ""))
                            
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
                                
                                st.cache_data.clear() 
                                for k in list(default_state.keys()):
                                    if k in st.session_state: del st.session_state[k]

                                time.sleep(1.5) 
                                st.session_state.app_mode = "END"
                                st.session_state.step = 1
                                st.rerun()

    elif st.session_state.app_mode == "END":
        df_all = load_universal_data().copy()
        
        if not df_all.empty and '상태' in df_all.columns:
            df_all['상태'] = df_all['상태'].fillna('').astype(str).str.strip()
            in_progress_df = df_all[df_all['상태'].str.contains('진행중', case=False, na=False)].copy()
        else:
            in_progress_df = pd.DataFrame()

        if step == 1:
            with st.container(border=True):
                st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 대상 LOT 선택</h4><br>", unsafe_allow_html=True)
                if in_progress_df.empty:
                    if not df_all.empty and '상태' in df_all.columns:
                        unique_status = df_all['상태'].unique()
                        st.info(f"현재 '진행중'인 작업이 없습니다. (현재 감지된 상태: {', '.join(unique_status)})")
                    else:
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
                        return f"LOT: {row.get('LOT NO.', '')} (시작: {row.get('시작시간', '')})"
                    
                    with sel_col2:
                        selected_id = st.selectbox("■ 마감할 LOT 선택", options, format_func=format_option)
                        st.session_state.target_unique_id = selected_id
                    
                    target_row = filtered_lots[filtered_lots['고유 ID'] == selected_id].iloc[0]
                    st.markdown(f"<div style='background-color: #FFC000; color: #000000; padding: 20px; border-radius: 10px; font-size: 1.2rem; font-weight: bold; box-shadow: 0 4px 10px rgba(0,0,0,0.1); margin-top: 15px;'>📌 모델명: {target_row.get('모델명(MI)', '')} &nbsp;|&nbsp; LOT: {target_row.get('LOT NO.', '')} &nbsp;|&nbsp; 시작시간: {target_row.get('시작시간', '')}</div>", unsafe_allow_html=True)

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
                            m, d = map(int, target_row.get('날짜', '').split('/'))
                            y = st.session_state.get("work_date", datetime.now()).year
                            s_date = datetime(y, m, d).date()
                            s_time = datetime.strptime(target_row.get('시작시간', '00:00'), "%H:%M").time()
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
                
                sbl_mod = st.session_state.get("model_name", "")
                if sbl_mod in ["MEM", "Centaur"]:
                    limit_y = st.session_state.sbl_limits.get(f'Yield_{sbl_mod}', 85.0)
                else:
                    limit_y = st.session_state.sbl_limits.get('Yield_Default', 85.0)
                
                limit_c = st.session_state.sbl_limits.get('Def_Comp', 10.0)
                limit_f = st.session_state.sbl_limits.get('Def_Front', 5.0)
                limit_r = st.session_state.sbl_limits.get('Def_Rear', 5.0)
                limit_o = st.session_state.sbl_limits.get('Def_Offset', 5.0)
                
                if comp_rate > limit_c and not st.session_state.get("comp_warned", False):
                    show_sbl_warning("완전불량", comp_rate, limit_c)
                    st.session_state.comp_warned = True
                elif front_rate > limit_f and not st.session_state.get("front_warned", False):
                    show_sbl_warning("전면불량", front_rate, limit_f)
                    st.session_state.front_warned = True
                elif rear_rate > limit_r and not st.session_state.get("rear_warned", False):
                    show_sbl_warning("배면불량", rear_rate, limit_r)
                    st.session_state.rear_warned = True
                elif offset_rate > limit_o and not st.session_state.get("offset_warned", False):
                    show_sbl_warning("옵셋불량", offset_rate, limit_o)
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
                with g_col1: st.plotly_chart(fig_donut, use_container_width=True, config={'displayModeBar': False})
                with g_col2: st.plotly_chart(fig_bar, use_container_width=True, config={'displayModeBar': False})

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
                                df_all = load_universal_data()
                                target_id = st.session_state.target_unique_id
                                df_target = df_all[df_all['고유 ID'] == target_id]
                                
                                if df_target.empty:
                                    st.error("오류: DB에서 해당 Lot를 찾을 수 없습니다.")
                                else:
                                    target_idx = df_target.index[0]
                                    target_row = df_target.iloc[0].to_dict()
                                    
                                    try:
                                        m, d = map(int, target_row.get('날짜', '').split('/'))
                                        y = st.session_state.get("work_date", datetime.now()).year
                                        s_date = datetime(y, m, d).date()
                                        s_time = datetime.strptime(target_row.get('시작시간', '00:00'), "%H:%M").time()
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
                                        for k in list(default_state.keys()):
                                            if k in st.session_state:
                                                del st.session_state[k]
                                                
                                        time.sleep(1.5)
                                        st.session_state.app_mode = "EDIT"
                                        st.session_state.step = 1
                                        st.rerun()

    # 💡 [데이터 수정] 모드 최적화 (비밀번호 및 내역 남기기)
    elif st.session_state.app_mode == "EDIT":
        with st.container(border=True):
            st.markdown("<h4 style='color: #1e293b; margin-top: 0; font-size: 1.1rem; border-bottom: 1px solid #cbd5e1; padding-bottom: 8px;'>■ 최근 저장 Data List</h4><br>", unsafe_allow_html=True)
            df_history = load_universal_data().copy()
            
            if not df_history.empty:
                cols_to_show = [c for c in EXCEL_COLUMNS if c in df_history.columns]
                if cols_to_show:
                    display_df = df_history[cols_to_show].copy()
                    display_df = display_df.iloc[::-1].head(20).copy()
                    
                    if not st.session_state.get("edit_unlocked", False):
                        # 💡 표를 먼저 보여주고 그 아래에 1:1 비율로 입력칸/버튼 배치
                        st.dataframe(display_df, hide_index=True, use_container_width=True) 
                        st.markdown("<hr style='border-color: #cbd5e1;'>", unsafe_allow_html=True)
                        st.info("🔒 데이터를 직접 수정하려면 아래에 관리자 비밀번호를 입력해주세요.")
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            edit_pwd = st.text_input("수정 비밀번호", type="password", label_visibility="collapsed", placeholder="비밀번호 입력 (6233)")
                        with c2:
                            if st.button("🔓 잠금 해제", use_container_width=True):
                                if edit_pwd == "6233":
                                    st.session_state.edit_unlocked = True
                                    st.rerun()
                                else:
                                    st.error("비밀번호가 일치하지 않습니다.")
                    else:
                        edited_df = st.data_editor(
                            display_df, 
                            use_container_width=True, 
                            hide_index=True,
                            column_config={"LOT NO.": st.column_config.TextColumn("LOT NO.")}
                        )
                        st.markdown("<br>", unsafe_allow_html=True)
                        edit_reason = st.text_input("📝 수정 사유 (수정 내역은 비고란에 자동으로 추가됩니다)", placeholder="예: 양품수량 오기입 수정")
                        
                        if st.button("Data 수정 적용", type="primary", use_container_width=True):
                            sheet = get_sheet()
                            changed = False
                            with st.spinner("구글 시트 업데이트 중..."):
                                for idx in edited_df.index:
                                    old_row = display_df.loc[idx].fillna("").astype(str).tolist()
                                    new_row = edited_df.loc[idx].fillna("").astype(str).tolist()
                                    if old_row != new_row:
                                        if edit_reason:
                                            memo_idx = EXCEL_COLUMNS.index('비고') if '비고' in EXCEL_COLUMNS else -1
                                            if memo_idx != -1:
                                                existing_memo = new_row[memo_idx]
                                                new_memo = f"{existing_memo} | [수정] {edit_reason}".strip(" |")
                                                new_row[memo_idx] = new_memo
                                        gspread_row = df_history.loc[idx, '_sheet_row']
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


# ==========================================
# 🚀 [라우팅 2] ADMINISTRATOR (관리자 패널 & 뷰어 진입)
# ==========================================
elif st.session_state.sys_menu == "admin":
    st.markdown("<h3 style='color:#1e293b; font-weight:900;'>⚙️ ADMINISTRATOR CONTROL PANEL</h3>", unsafe_allow_html=True)
    
    config = load_shared_config() or {}
    df = load_universal_data()
    all_models = df['모델명(MI)'].dropna().unique().tolist() if not df.empty else ["ALL_MODELS"]
    
    # 💡 모델 중복 선택 방지 로직 (Dynamic Options)
    opt_std = [m for m in all_models if m not in st.session_state.sel_inc]
    opt_inc = [m for m in all_models if m not in st.session_state.sel_std]
    
    with st.form("admin_config_form"):
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**1. 모니터링 대상 모델 선택 (중복 불가)**")
            def update_std(): pass
            def update_inc(): pass
            sel_std = st.multiselect("기본 양품율 적용 모델 (Yield 1)", opt_std, key="sel_std", on_change=update_std)
            sel_inc = st.multiselect("전/배 포함 양품율 적용 모델 (Yield 2)", opt_inc, key="sel_inc", on_change=update_inc)
            
            st.markdown("<br>**2. 뷰어 기본 설정**", unsafe_allow_html=True)
            time_range = st.selectbox("기본 조회 기간 (Default Time Range)", ["6H", "24H", "48H", "72H", "96H"], index=["6H", "24H", "48H", "72H", "96H"].index(config.get("time_range", "48H")))
            auto_rotate = st.checkbox("자동 로테이션 활성화 (10분 단위로 선택된 모델 순환 표출)", value=config.get("auto_rotate_active", False))
            
        with col2:
            st.markdown("**3. SBL (Sub-Block Limit) 알람 임계치 설정 (%)**")
            sbl_limits = config.get("sbl_limits", {})
            
            # 💡 각 모델별 양품율 SBL 세분화 적용
            sbl_yield_def = st.number_input("📉 양품율 SBL (기본)", value=float(sbl_limits.get("Yield_Default", 85.0)), step=0.1)
            sbl_yield_cen = st.number_input("📉 양품율 SBL (Centaur)", value=float(sbl_limits.get("Yield_Centaur", 91.4)), step=0.1)
            sbl_yield_mem = st.number_input("📉 양품율 SBL (MEM)", value=float(sbl_limits.get("Yield_MEM", 93.2)), step=0.1)
            
            sbl_comp = st.number_input("📈 완전불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Comp", 10.0)), step=0.1)
            sbl_front = st.number_input("📈 전면불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Front", 5.0)), step=0.1)
            sbl_rear = st.number_input("📈 배면불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Rear", 5.0)), step=0.1)
            sbl_offset = st.number_input("📈 옵셋불량 SBL (이상일 때 알람)", value=float(sbl_limits.get("Def_Offset", 5.0)), step=0.1)

        st.markdown("<br>", unsafe_allow_html=True)
        if st.form_submit_button("💾 설정 저장 및 대시보드 클라우드 반영 (Save to Cloud)", use_container_width=True, type="primary"):
            new_config = {
                "sel_std": sel_std,
                "sel_inc": sel_inc,
                "time_range": time_range,
                "auto_rotate_active": auto_rotate,
                "sbl_limits": {
                    "Yield_Default": sbl_yield_def,
                    "Yield_Centaur": sbl_yield_cen,
                    "Yield_MEM": sbl_yield_mem,
                    "Def_Comp": sbl_comp,
                    "Def_Front": sbl_front,
                    "Def_Rear": sbl_rear,
                    "Def_Offset": sbl_offset
                },
                "model_color_dict": config.get("model_color_dict", {})
            }
            if save_shared_config(new_config):
                st.cache_data.clear()
                time.sleep(1.0)
                st.rerun()

    st.markdown("<hr style='border-color: #cbd5e1; margin-top: 30px; margin-bottom: 30px;'>", unsafe_allow_html=True)
    st.info("💡 위에서 설정값을 저장한 후, 아래 버튼을 눌러 모니터링 뷰어 화면을 확인할 수 있습니다.")
    
    if st.button("👁️ 설정된 VIEWER 화면 실행하기", type="primary", use_container_width=True):
        st.session_state.sys_menu = "viewer"
        st.rerun()


# ==========================================
# 🚀 [라우팅 3] VIEWER 모드 (ADMIN에서 진입)
# ==========================================
elif st.session_state.sys_menu == "viewer":
    config = load_shared_config()
    if not config:
        st.warning("📡 설정값이 없습니다. ADMINISTRATOR 메뉴에서 설정을 완료하세요.")
        st.stop()
    if "viewer_time_range" not in st.session_state:
        st.session_state.viewer_time_range = config.get("time_range", "48H")

    col1, col2 = st.columns([0.65, 0.35])
    with col1:
        st.markdown(f"<div class='command-header' style='font-size: 1.8rem; margin-top: 5px;'><span class='live-dot'></span>VISION DATA KEY-IN SYSTEM (VIEWER)</div>", unsafe_allow_html=True)
        st.markdown("<div style='color: #10b981; font-size: 0.85rem; margin-bottom: 15px; font-weight:bold;'>Shared Dashboard (View Only)</div>", unsafe_allow_html=True)
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        vc1, vc2, vc3 = st.columns(3)
        with vc1:
            if st.button("⬅️ ADMIN", use_container_width=True):
                st.session_state.sys_menu = "admin"
                st.rerun()
        with vc2:
            if st.button("🔄 Rotate", use_container_width=True, key="viewer_manual_rotate"):
                st.session_state.rotate_idx += 1
                st.rerun()
        with vc3:
            if st.button("RELOAD", type="primary", use_container_width=True, key="viewer_reload"):
                st.cache_data.clear()
                st.rerun()

    time_range = st.session_state.viewer_time_range
    now_kst = datetime.now(timezone(timedelta(hours=9))).replace(tzinfo=None)
    target_end_date = now_kst.date() 

    df = load_universal_data().copy()
    if df.empty: 
        st.warning("데이터베이스에 렌더링할 정보가 전혀 없습니다.")
        st.stop()
        
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

        with st.container(border=True): st.plotly_chart(make_donut_chart(f"OVERALL ({time_range})", o_t, o_g, o_c, o_f, o_r, o_o), use_container_width=True, config={'displayModeBar': False})
        with st.container(border=True): st.plotly_chart(make_donut_chart("YESTERDAY", y_t, y_g, y_c, y_f, y_r, y_o), use_container_width=True, config={'displayModeBar': False})
        with st.container(border=True): st.plotly_chart(make_donut_chart("LAST 6 HOURS", h_t, h_g, h_c, h_f, h_r, h_o), use_container_width=True, config={'displayModeBar': False})

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
            st.plotly_chart(fig_yld, use_container_width=True, config={'displayModeBar': False})
            
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
            st.plotly_chart(fig_def, use_container_width=True, config={'displayModeBar': False})

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
# 🚀 [라우팅 4] EXIT (시스템 안전 종료)
# ==========================================
elif st.session_state.sys_menu == "exit":
    st.markdown("<br><br><br><br>", unsafe_allow_html=True)
    exit_script = """
    <script>
    setTimeout(function() {
        try { window.parent.close(); } catch(e) {}
        window.parent.document.body.innerHTML = `
            <div style="display:flex; justify-content:center; align-items:center; height:100vh; background-color:#f1f5f9; flex-direction:column;">
                <h1 style="color:#1e293b; font-family:sans-serif; font-size:3rem; margin-bottom:10px;">VISION DATA KEY-IN SYSTEM</h1>
                <h2 style="color:#ef4444; font-family:sans-serif; font-size:2rem; margin-bottom:20px;">시스템이 안전하게 종료되었습니다.</h2>
                <p style="color:#64748b; font-family:sans-serif; font-size:1.2rem; font-weight:bold;">보안을 위해 현재 열려있는 브라우저 창(탭)을 닫아주세요.</p>
            </div>
        `;
    }, 500);
    </script>
    """
    components.html(exit_script, height=0, width=0)

# ==========================================
# 🛡️ 최하단: 화면 렌더링 완료 후 깜빡임 없는 스크립트 실행 (Flashing 방지)
# ==========================================
bottom_js = f"""
<script>
const pDoc = window.parent.document;
if (pDoc) {{
    // 💡 1. 커스텀 플로팅 토글 및 EXIT 버튼 설정
    if ("{st.session_state.sys_menu}" === "keyin" && "{st.session_state.unlocked}" === "True") {{
        let toggleBtn = pDoc.getElementById('custom-sidebar-toggle');
        if (!toggleBtn) {{
            toggleBtn = pDoc.createElement('div');
            toggleBtn.id = 'custom-sidebar-toggle';
            toggleBtn.innerHTML = '<span style="font-size:1.4rem; line-height:1;">☰</span> <span style="margin-top:2px;">사이드바 토글</span>';
            toggleBtn.style.cssText = 'position:fixed; top:20px; left:20px; z-index:999999; background:#1e293b; color:#ffffff; padding:10px 15px; border-radius:8px; cursor:pointer; font-weight:bold; box-shadow:0 4px 10px rgba(0,0,0,0.3); display:flex; align-items:center; gap:8px; border:2px solid #cbd5e1; transition:all 0.2s; font-family:sans-serif;';
            toggleBtn.onmouseover = () => {{ toggleBtn.style.background = '#3b82f6'; toggleBtn.style.borderColor = '#ffffff'; }};
            toggleBtn.onmouseout = () => {{ toggleBtn.style.background = '#1e293b'; toggleBtn.style.borderColor = '#cbd5e1'; }};
            pDoc.body.appendChild(toggleBtn);
        }}
        toggleBtn.style.display = 'flex';
        toggleBtn.onclick = function() {{
            const expandDiv = pDoc.querySelector('[data-testid="collapsedControl"]');
            const collapseDiv = pDoc.querySelector('[data-testid="stSidebarCollapseButton"]');
            if (collapseDiv && collapseDiv.getBoundingClientRect().width > 0) {{
                const btn = collapseDiv.querySelector('button') || collapseDiv;
                btn.click();
            }} else if (expandDiv) {{
                const btn = expandDiv.querySelector('button') || expandDiv;
                btn.click();
            }}
        }};
        
        let exitBtn = pDoc.getElementById('custom-exit-toggle');
        if (!exitBtn) {{
            exitBtn = pDoc.createElement('div');
            exitBtn.id = 'custom-exit-toggle';
            exitBtn.innerHTML = '<span style="font-size:1.4rem; line-height:1;">🚪</span> <span style="margin-top:2px;">EXIT</span>';
            exitBtn.style.cssText = 'position:fixed; top:20px; right:20px; z-index:999999; background:#ef4444; color:#ffffff; padding:10px 15px; border-radius:8px; cursor:pointer; font-weight:bold; box-shadow:0 4px 10px rgba(0,0,0,0.3); display:flex; align-items:center; gap:8px; border:2px solid #fca5a5; transition:all 0.2s; font-family:sans-serif;';
            exitBtn.onmouseover = () => {{ exitBtn.style.background = '#dc2626'; }};
            exitBtn.onmouseout = () => {{ exitBtn.style.background = '#ef4444'; }};
            exitBtn.onclick = function() {{
                const btns = pDoc.querySelectorAll('button');
                for(let b of btns) {{ if(b.innerText.includes('HIDDEN_EXIT_TRIGGER')) {{ b.click(); break; }} }}
            }};
            pDoc.body.appendChild(exitBtn);
        }}
        exitBtn.style.display = 'flex';

    }} else {{
        let toggleBtn = pDoc.getElementById('custom-sidebar-toggle');
        let exitBtn = pDoc.getElementById('custom-exit-toggle');
        if (toggleBtn) toggleBtn.style.display = 'none';
        if (exitBtn) exitBtn.style.display = 'none';
    }}

    // 💡 2. 뷰어 화면 오토 로테이션 및 리로드 연동
    if ("{st.session_state.sys_menu}" === "viewer") {{
        setTimeout(function() {{
            const btns = pDoc.querySelectorAll('button');
            for(let i=0; i<btns.length; i++){{ if(btns[i].textContent && btns[i].textContent.includes('RELOAD')){{ btns[i].click(); break; }} }}
        }}, 1800000); 
        {'setTimeout(function() { const btns = pDoc.querySelectorAll("button"); for(let i=0; i<btns.length; i++){ if(btns[i].textContent && btns[i].textContent.includes("Manual Rotate")){ btns[i].click(); break; } } }, 600000);' if st.session_state.get('auto_rotate_active', False) else ''}
    }}

    // 💡 3. Streamlit Cloud UI 철저한 강제 삭제 (DOM Nuke) - 깜빡임 원천 차단
    const nukeNode = (el) => {{ if(el && el.parentNode) el.parentNode.removeChild(el); }};
    const destroyStreamlitUI = () => {{
        let docs = [document];
        try {{ if (window.parent && window.parent.document) docs.push(window.parent.document); }} catch(e){{}}
        try {{ if (window.top && window.top.document && window.top !== window.parent) docs.push(window.top.document); }} catch(e){{}}
        
        docs.forEach(doc => {{
            try {{
                // iframe 및 배지 DOM 자체를 뜯어내서 삭제
                doc.querySelectorAll('iframe').forEach(f => {{ if(f.src && (f.src.includes('badge') || f.title.includes('Toolbar'))) nukeNode(f); }});
                doc.querySelectorAll('[data-testid="manage-app-button"], [data-testid="stAppDeployButton"], .stDeployButton, div[class^="viewerBadge"]').forEach(nukeNode);
                
                doc.querySelectorAll('div, a, button, span').forEach(el => {{
                    if (el.textContent && (el.textContent === '< Manage app' || el.textContent === 'Manage app' || el.textContent.includes('View profile'))) {{
                        el.style.setProperty('display', 'none', 'important');
                        if (el.parentElement) el.parentElement.style.setProperty('display', 'none', 'important');
                    }}
                }});
            }} catch(e) {{}}
        }});
    }};
    destroyStreamlitUI(); 
    setInterval(destroyStreamlitUI, 50); 
}}
</script>
"""
components.html(bottom_js, height=0, width=0)
