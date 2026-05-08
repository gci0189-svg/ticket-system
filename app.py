"""
票務系統 v2.0
流程：上傳 xlsx → 預覽所有工作表 → 選一張 → 確認解析規則 → 產生標籤（四場次分組）→ 複製 → 標記已列印
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
import gspread
from google.oauth2.service_account import Credentials

# ══════════════════════════════════════════════════════════
# 頁面設定
# ══════════════════════════════════════════════════════════
st.set_page_config(page_title="票務系統", page_icon="🎫", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&family=Space+Mono:wght@400;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans TC', sans-serif; }

.hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #0f3460 100%);
    border-radius: 16px; padding: 1.5rem 2rem; margin-bottom: 1.5rem;
    border: 1px solid rgba(255,255,255,0.08);
}
.hero h1 { font-family:'Space Mono',monospace; color:#e94560; font-size:1.8rem; margin:0 0 0.2rem; }
.hero p  { color:rgba(255,255,255,0.5); margin:0; font-size:0.85rem; }

.step-box {
    border:2px solid #e94560; border-radius:12px;
    padding:1.2rem 1.5rem; margin-bottom:1.2rem;
}
.step-title { color:#e94560; font-weight:700; font-size:1rem; margin-bottom:0.75rem; }

.rule-box {
    background:#f8f9fa; border:1px solid #dee2e6; border-radius:8px;
    padding:0.8rem 1rem; margin-bottom:0.5rem; font-size:0.88rem;
}
.rule-key { color:#888; font-size:0.78rem; margin-bottom:0.15rem; }
.rule-val { font-weight:600; color:#1a1a2e; }

.session-header {
    background:#1a1a2e; color:#e94560;
    font-family:'Space Mono',monospace; font-size:0.85rem; font-weight:700;
    padding:0.6rem 1rem; border-radius:8px 8px 0 0; margin-top:1.2rem;
    display:flex; justify-content:space-between; align-items:center;
}
.stat-box {
    background:linear-gradient(135deg,#e94560,#c73652);
    color:white; border-radius:10px; padding:0.8rem 1rem; text-align:center;
}
.stat-num { font-family:'Space Mono',monospace; font-size:1.8rem; font-weight:700; line-height:1; }
.stat-label { font-size:0.75rem; opacity:0.85; margin-top:0.2rem; }

.warn-box {
    background:#fff3cd; border:1px solid #ffc107; border-radius:8px;
    padding:0.8rem 1rem; margin-bottom:0.8rem; font-size:0.88rem;
}
.ok-box {
    background:#d4edda; border:1px solid #28a745; border-radius:8px;
    padding:0.8rem 1rem; margin-bottom:0.8rem; font-size:0.88rem;
}

#MainMenu, footer { visibility:hidden; }
.stDeployButton { display:none; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# Google Sheets
# ══════════════════════════════════════════════════════════
SCOPES = ["https://www.googleapis.com/auth/spreadsheets",
          "https://www.googleapis.com/auth/drive.readonly"]

@st.cache_resource(ttl=300)
def get_gc():
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=SCOPES)
        return gspread.authorize(creds)
    except:
        return None

def load_history(sid: str) -> set:
    gc = get_gc()
    if not gc or not sid:
        return set()
    try:
        sh = gc.open_by_key(sid)
        try:
            ws = sh.worksheet("已列印紀錄")
        except:
            ws = sh.add_worksheet("已列印紀錄", 1000, 3)
            ws.append_row(["唯一識別碼", "列印時間", "標籤內容"])
            return set()
        vals = ws.col_values(1)
        return {v for v in vals if v and v != "唯一識別碼"}
    except Exception as e:
        st.warning(f"載入歷史失敗：{e}")
        return set()

def save_history(sid: str, keys: list, labels: list) -> bool:
    gc = get_gc()
    if not gc:
        st.error("Google Sheets 連線失敗")
        return False
    try:
        sh = gc.open_by_key(sid)
        try:
            ws = sh.worksheet("已列印紀錄")
        except:
            ws = sh.add_worksheet("已列印紀錄", 1000, 3)
            ws.append_row(["唯一識別碼", "列印時間", "標籤內容"])
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        ws.append_rows([[k, now, l] for k, l in zip(keys, labels)])
        return True
    except Exception as e:
        st.error(f"儲存失敗：{e}")
        return False


# ══════════════════════════════════════════════════════════
# 解析函數：LINE 會員格式（5月親子｜會員購票）
#
# 欄位對應（已與使用者確認）：
#   Row0 = 標題列
#   Row1 = 彙總計算列（跳過）
#   B欄(index=1) = 編號
#   D欄(index=3) = 姓名
#   G欄(index=6) = 日期場次
#   H欄(index=7) = 座位
#   I欄(index=8) = 張數
#   J欄(index=9) = 票價
#
# 列印條件：座位、張數、票價三欄都有值
# 合併：相同「編號」的多行 → 張數加總、取最早場次
# 唯一識別碼：{編號}_{工作表名稱}
# 標籤格式：NO.{編號} {最早場次} 貴賓｜{姓名} X {總張數}
# ══════════════════════════════════════════════════════════

def parse_session(raw: str):
    """解析日期場次字串，回傳 (排序key, 顯示文字)"""
    if not raw:
        return (99, 99, 99), ""
    raw = str(raw)
    m = re.search(r'(\d{1,2})[/／\-月](\d{1,2})', raw)
    month = int(m.group(1)) if m else 99
    day   = int(m.group(2)) if m else 99
    if any(x in raw for x in ["14:30","15:00","下午","PM","pm"]):
        sess_str, sess_ord = "下午", 1
    elif any(x in raw for x in ["10:30","10:00","上午","AM","am"]):
        sess_str, sess_ord = "上午", 0
    else:
        sess_str, sess_ord = "", 0
    display = f"{month}/{day} {sess_str}".strip()
    return (month, day, sess_ord), display


def parse_member_sheet(df: pd.DataFrame, sheet_name: str, history_set: set):
    """
    解析 LINE 會員格式工作表。
    回傳 (tickets, skipped, warnings)
    """
    rows = df.values.tolist()

    COL_ID    = 1   # B：編號
    COL_NAME  = 3   # D：姓名
    COL_DATE  = 6   # G：日期場次
    COL_SEAT  = 7   # H：座位
    COL_COUNT = 8   # I：張數
    COL_PRICE = 9   # J：票價

    def get(row, col):
        if col >= len(row):
            return ""
        v = row[col]
        return str(v).strip() if v is not None else ""

    merged   = {}   # unique_key → dict
    warnings = []
    last_id  = None  # 追蹤最近一個有效編號

    for i, row in enumerate(rows):
        if i < 2:
            continue  # 跳過 Row0（標題）和 Row1（彙總）

        seat      = get(row, COL_SEAT)
        count_raw = get(row, COL_COUNT)
        price     = get(row, COL_PRICE)

        # 列印條件：三欄都要有值
        if not seat or not count_raw or not price:
            continue

        # 張數解析
        count_clean = re.sub(r'[^0-9]', '', count_raw)
        if not count_clean:
            warnings.append(f"第 {i+1} 行：張數無法解析（{count_raw!r}）")
            continue
        count = int(count_clean)
        if count <= 0:
            continue

        # 票價必須像數字
        if not re.search(r'\d', price):
            continue

        row_id   = get(row, COL_ID)
        name     = get(row, COL_NAME)
        date_raw = get(row, COL_DATE)

        # 更新 last_id
        id_digits = re.sub(r'[^0-9]', '', row_id)
        if id_digits:
            last_id = id_digits
        
        if not last_id:
            warnings.append(f"第 {i+1} 行：找不到編號，已略過")
            continue

        key = f"{last_id}_{sheet_name}"
        sort_key, display = parse_session(date_raw)

        if key not in merged:
            merged[key] = {
                "key":              key,
                "id":               last_id,
                "name":             name or "(未填姓名)",
                "sheet":            sheet_name,
                "total_count":      0,
                "earliest_sort":    sort_key,
                "earliest_display": display,
                "is_new":           key not in history_set,
            }
        else:
            # 更新最早場次
            if sort_key < merged[key]["earliest_sort"]:
                merged[key]["earliest_sort"]   = sort_key
                merged[key]["earliest_display"] = display
            # 補姓名（跨行的附屬行可能是空的）
            if (not merged[key]["name"] or merged[key]["name"] == "(未填姓名)") and name:
                merged[key]["name"] = name

        merged[key]["total_count"] += count

    # 產生標籤
    tickets = []
    skipped = []
    for key, info in merged.items():
        label = (f"NO.{info['id']} {info['earliest_display']} "
                 f"貴賓｜{info['name']} X {info['total_count']}")
        entry = {**info, "label": label, "count": info["total_count"]}
        if info["is_new"]:
            tickets.append(entry)
        else:
            skipped.append(entry)

    tickets.sort(key=lambda x: x["earliest_sort"])
    skipped.sort(key=lambda x: x["earliest_sort"])
    return tickets, skipped, warnings


def group_by_session(tickets: list) -> dict:
    grouped = {}
    for t in tickets:
        d = t["earliest_display"]
        grouped.setdefault(d, []).append(t)
    return dict(sorted(grouped.items(), key=lambda x: x[1][0]["earliest_sort"]))


# ══════════════════════════════════════════════════════════
# Session State
# ══════════════════════════════════════════════════════════
defaults = {
    "history_sid":    "",
    "history_set":    set(),
    "raw_sheets":     {},
    "selected_sheet": None,
    "rule_confirmed": False,
    "tickets":        [],
    "skipped":        [],
    "warnings":       [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# ══════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════
st.markdown("""
<div class="hero">
  <h1>🎫 票務系統</h1>
  <p>上傳報名 xlsx → 確認規則 → 產生信封標籤 → 複製貼到標籤機</p>
</div>
""", unsafe_allow_html=True)

# ── 側邊欄 ────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ 歷史紀錄設定")
    sid_input = st.text_input(
        "Google Sheets ID",
        value=st.session_state.history_sid,
        placeholder="貼上試算表 ID",
        help="記錄已列印資料，防止重複列印"
    )
    if sid_input != st.session_state.history_sid:
        st.session_state.history_sid = sid_input
        st.session_state.history_set = set()

    if sid_input:
        if st.button("🔄 重新載入歷史", use_container_width=True):
            with st.spinner("連線中..."):
                st.session_state.history_set = load_history(sid_input)
            st.success(f"已載入 {len(st.session_state.history_set)} 筆")
        elif not st.session_state.history_set:
            st.session_state.history_set = load_history(sid_input)

    st.divider()
    st.markdown(f"📦 歷史紀錄：**{len(st.session_state.history_set)}** 筆")
    st.divider()
    if st.button("🔄 重新開始", use_container_width=True):
        for k in ["raw_sheets","selected_sheet","rule_confirmed","tickets","skipped","warnings"]:
            st.session_state[k] = defaults[k]
        st.rerun()


# ══════════════════════════════════════════════════════════
# STEP 1：上傳檔案
# ══════════════════════════════════════════════════════════
st.markdown('<div class="step-box">', unsafe_allow_html=True)
st.markdown('<div class="step-title">STEP 1 ｜ 上傳 xlsx 報名表</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("選取或拖曳 xlsx 檔案", type=["xlsx"], label_visibility="collapsed")
if uploaded:
    try:
        xls = pd.ExcelFile(BytesIO(uploaded.read()))
        raw = {}
        for sname in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sname, header=None, dtype=str)
            raw[sname] = df.fillna("")
        if raw != st.session_state.raw_sheets:
            st.session_state.raw_sheets     = raw
            st.session_state.selected_sheet = None
            st.session_state.rule_confirmed  = False
            st.session_state.tickets         = []
            st.session_state.skipped         = []
            st.session_state.warnings        = []
        st.success(f"✅ 已載入：{uploaded.name}，共 {len(raw)} 個工作表")
    except Exception as e:
        st.error(f"讀取失敗：{e}")

st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# STEP 2：預覽 + 選擇工作表
# ══════════════════════════════════════════════════════════
if st.session_state.raw_sheets:
    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.markdown('<div class="step-title">STEP 2 ｜ 預覽工作表，選擇要處理的那一張</div>', unsafe_allow_html=True)

    sheet_names = list(st.session_state.raw_sheets.keys())
    tabs = st.tabs(sheet_names)
    for tab, sname in zip(tabs, sheet_names):
        with tab:
            preview = st.session_state.raw_sheets[sname].iloc[:8, :12]
            st.dataframe(preview, use_container_width=True, height=230)

    st.markdown("<br>", unsafe_allow_html=True)
    cur_idx = (sheet_names.index(st.session_state.selected_sheet) + 1
               if st.session_state.selected_sheet in sheet_names else 0)
    selected = st.selectbox(
        "選擇要產生標籤的工作表",
        options=["（請選擇）"] + sheet_names,
        index=cur_idx
    )
    if selected != "（請選擇）" and selected != st.session_state.selected_sheet:
        st.session_state.selected_sheet = selected
        st.session_state.rule_confirmed  = False
        st.session_state.tickets         = []
        st.session_state.skipped         = []
        st.session_state.warnings        = []
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# STEP 3：確認解析規則
# ══════════════════════════════════════════════════════════
if st.session_state.selected_sheet:
    sname = st.session_state.selected_sheet
    df    = st.session_state.raw_sheets[sname]

    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.markdown(f'<div class="step-title">STEP 3 ｜ 確認解析規則：{sname}</div>', unsafe_allow_html=True)

    row0_str = " ".join(str(c) for c in df.iloc[0].tolist()) if len(df) > 0 else ""
    is_member = ("姓名" in row0_str and "張數" in row0_str and "座位" in row0_str)

    if is_member:
        st.markdown('<div class="ok-box">✅ 偵測到：<strong>LINE 會員格式</strong></div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            for key, val in [
                ("標題列位置", "第 1 行（Row 1）"),
                ("跳過",       "第 2 行（彙總計算列）"),
                ("列印條件",   "座位 ＋ 張數 ＋ 票價 三欄都有值"),
                ("合併邏輯",   "相同編號的多行 → 張數加總，取最早場次"),
            ]:
                st.markdown(f'<div class="rule-box"><div class="rule-key">{key}</div><div class="rule-val">{val}</div></div>', unsafe_allow_html=True)
        with col2:
            for key, val in [
                ("唯一識別碼", "編號 ＋ 工作表名稱"),
                ("標籤格式",   "NO.{編號} {最早場次} 貴賓｜{姓名} X {總張數}"),
                ("排序方式",   "依最早場次（日期 → 上午 → 下午）"),
                ("輸出分組",   "每個場次獨立一區，各自可複製"),
            ]:
                st.markdown(f'<div class="rule-box"><div class="rule-key">{key}</div><div class="rule-val">{val}</div></div>', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)
        col_a, col_b = st.columns([3, 1])
        with col_a:
            st.info("⚠️ 請確認以上規則符合這份 xlsx 的格式，再按右側按鈕")
        with col_b:
            if st.button("✅ 確認，開始產生標籤", type="primary", use_container_width=True):
                with st.spinner("解析中..."):
                    t, s, w = parse_member_sheet(df, sname, st.session_state.history_set)
                st.session_state.tickets        = t
                st.session_state.skipped        = s
                st.session_state.warnings       = w
                st.session_state.rule_confirmed = True
                st.rerun()
    else:
        st.markdown("""
        <div class="warn-box">
        ⚠️ <strong>這張工作表的格式目前尚未設定解析規則。</strong><br><br>
        請把這份 xlsx 傳給管理員確認格式後更新程式，才能安全產生標籤。<br>
        <strong>現場取票時間寶貴，絕對不能猜測格式。</strong>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# STEP 4：標籤結果
# ══════════════════════════════════════════════════════════
if st.session_state.rule_confirmed:
    tickets  = st.session_state.tickets
    skipped  = st.session_state.skipped
    warnings = st.session_state.warnings

    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.markdown('<div class="step-title">STEP 4 ｜ 標籤結果</div>', unsafe_allow_html=True)

    if warnings:
        with st.expander(f"⚠️ {len(warnings)} 筆資料需要注意", expanded=True):
            for w in warnings:
                st.markdown(f"- {w}")

    # 統計
    total_count  = sum(t["count"] for t in tickets)
    session_count = len(set(t["earliest_display"] for t in tickets))
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="stat-box"><div class="stat-num">{len(tickets)}</div><div class="stat-label">待列印（人）</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="stat-box" style="background:linear-gradient(135deg,#1565c0,#0d47a1)"><div class="stat-num">{total_count}</div><div class="stat-label">總張數</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="stat-box" style="background:linear-gradient(135deg,#2e7d32,#1b5e20)"><div class="stat-num">{session_count}</div><div class="stat-label">場次數</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="stat-box" style="background:linear-gradient(135deg,#555,#333)"><div class="stat-num">{len(skipped)}</div><div class="stat-label">略過（已列印）</div></div>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if not tickets:
        st.info("沒有新的待列印資料（全部都已列印過）")
    else:
        grouped = group_by_session(tickets)

        st.markdown("#### 📋 待列印標籤（依場次分組）")
        st.caption("點入文字框 → Ctrl+A 全選 → Ctrl+C 複製 → 貼到標籤機軟體")

        for session_display, sess_tickets in grouped.items():
            sess_people = len(sess_tickets)
            sess_total  = sum(t["count"] for t in sess_tickets)
            label_text  = "\n".join(t["label"] for t in sess_tickets)

            st.markdown(
                f'<div class="session-header">'
                f'<span>🗓 {session_display}</span>'
                f'<span>{sess_people} 人 ／ {sess_total} 張</span>'
                f'</div>',
                unsafe_allow_html=True
            )
            st.text_area(
                label=session_display,
                value=label_text,
                height=min(400, max(100, sess_people * 33)),
                key=f"ta_{session_display}",
                label_visibility="collapsed"
            )

        # 全部合併
        st.markdown("---")
        st.markdown("#### 📄 全部場次合併（如需一次複製全部）")
        all_text = "\n".join(t["label"] for t in tickets)
        st.text_area(
            "全部",
            value=all_text,
            height=min(500, max(150, len(tickets) * 33)),
            label_visibility="collapsed"
        )

        # 標記已列印
        st.markdown("---")
        st.markdown("#### ✅ 列印完成後請點此歸檔")
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            st.info(f"點「標記為已列印」後，這 **{len(tickets)}** 筆會寫入 Google Sheets，下次不再重複出現。")
        with col_p2:
            if not st.session_state.history_sid:
                st.warning("請先在左側設定 Google Sheets ID")
            else:
                if st.button("✅ 標記為已列印", type="primary", use_container_width=True):
                    keys   = [t["key"]   for t in tickets]
                    labels = [t["label"] for t in tickets]
                    if save_history(st.session_state.history_sid, keys, labels):
                        st.session_state.history_set.update(keys)
                        st.session_state.tickets        = []
                        st.session_state.skipped        = []
                        st.session_state.rule_confirmed = False
                        st.success(f"✅ 已歸檔 {len(keys)} 筆！")
                        st.rerun()

    if skipped:
        with st.expander(f"⏭️ 略過 {len(skipped)} 筆（歷史紀錄中已列印過）"):
            for t in skipped:
                st.markdown(f'<span style="color:#aaa;font-size:0.83rem;">{t["label"]}</span>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)
