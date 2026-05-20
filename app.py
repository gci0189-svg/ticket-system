"""
票務系統 v2.0
流程：上傳 xlsx → 預覽所有工作表 → 選一張 → 確認解析規則 → 產生標籤（四場次分組）→ 複製 → 標記已列印
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
from io import BytesIO
from collections import defaultdict
import gspread
from google.oauth2.service_account import Credentials
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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

HISTORY_SHEET_ID = "1McLgRi-4haGs1orOXlCaucL6lM9E9uFgt4El8R4ZdkA"

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


# ══════════════════════════════════════════════════════════
# 簽到表產生函數：LINE 會員格式
# ══════════════════════════════════════════════════════════

def _cjk_len(s):
    """計算含中文的字串顯示寬度"""
    return sum(2 if "\u4e00" <= c <= "\u9fff" or "\u3000" <= c <= "\u303f" or "\uff00" <= c <= "\uffef" else 1 for c in str(s))

def _thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def _has_value(row, idx):
    if idx >= len(row): return False
    v = row[idx]
    if v is None: return False
    if isinstance(v, (int, float)): return v != 0
    return str(v).strip() != ""

def _get(row, idx):
    if idx >= len(row): return ""
    v = row[idx]
    if v is None: return ""
    if isinstance(v, datetime): return v.strftime("%Y/%m/%d")
    return str(v).strip()

def generate_signin_excel(file_bytes: bytes, sheet_name: str, show_name: str) -> bytes:
    """
    從報名表 xlsx 產生簽到表 Excel。
    目前支援 LINE 會員格式（5月親子｜會員購票）。
    回傳 Excel bytes。
    """
    wb_src = load_workbook(BytesIO(file_bytes), read_only=True)
    ws_src = wb_src[sheet_name]
    rows = list(ws_src.iter_rows(values_only=True))

    COL_ID=1; COL_SNS=2; COL_NAME=3; COL_TEL=4
    COL_DATE=6; COL_SEAT=7; COL_COUNT=8; COL_PRICE=9

    merged = {}
    last_id = None
    for i, row in enumerate(rows):
        if i < 2: continue
        if not _has_value(row, COL_SEAT): continue
        if not _has_value(row, COL_COUNT): continue
        if not _has_value(row, COL_PRICE): continue

        count_raw = _get(row, COL_COUNT)
        count_clean = re.sub(r"[^0-9]", "", count_raw)
        if not count_clean: continue
        count = int(count_clean)
        if count <= 0: continue

        row_id = _get(row, COL_ID)
        id_digits = re.sub(r"[^0-9]", "", row_id)
        if id_digits: last_id = id_digits
        if not last_id: continue

        name = _get(row, COL_NAME)
        sns  = _get(row, COL_SNS)
        tel  = _get(row, COL_TEL)
        date_raw = _get(row, COL_DATE)

        if last_id not in merged:
            m = re.search(r"\d{4}[/\-](\d{1,2})[/\-](\d{1,2})", date_raw)
            month = int(m.group(1)) if m else 99
            day   = int(m.group(2)) if m else 99
            sess_ord = 1 if any(x in date_raw for x in ["14:30","下午"]) else 0
            sess_str = "下午" if sess_ord else "上午"
            wday_m = re.search(r"[(（](.)[ )）]", date_raw)
            wday = wday_m.group(1) if wday_m else ""
            time_str = "14:30" if sess_ord else "10:30"
            full_date = f"2026/{month:02d}/{day:02d}（{wday}）{time_str}"
            merged[last_id] = {
                "id": last_id, "name": name, "sns": sns, "tel": tel,
                "seats": [], "total": 0,
                "sort": (month, day, sess_ord),
                "session": f"{month}/{day} {sess_str}",
                "full_date": full_date
            }
        else:
            if not merged[last_id]["name"] and name: merged[last_id]["name"] = name
            if not merged[last_id]["sns"] and sns:   merged[last_id]["sns"] = sns
            if not merged[last_id]["tel"] and tel:   merged[last_id]["tel"] = tel

        seat = _get(row, COL_SEAT)
        if seat:
            for s in re.split(r"[\n\r]+", seat):
                s = s.strip()
                if s: merged[last_id]["seats"].append(s)
        merged[last_id]["total"] += count

    sorted_data = sorted(merged.values(), key=lambda x: (x["sort"], int(x["id"])))
    sessions = defaultdict(list)
    for r in sorted_data:
        sessions[r["session"]].append(r)

    # 樣式
    HDR_FONT = Font(name="Arial", bold=True, size=11)
    SUB_FILL = PatternFill("solid", start_color="E94560")
    SUB_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
    BODY_FONT = Font(name="Arial", size=10)
    BOLD_FONT = Font(name="Arial", bold=True, size=10)
    ALT_FILL  = PatternFill("solid", start_color="FFF5F7")
    TOT_FILL  = PatternFill("solid", start_color="EEEEEE")
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    wb_out = Workbook()
    wb_out.remove(wb_out.active)

    for session_display, sess_rows in sessions.items():
        safe_name = session_display.replace("/", "-")
        ws = wb_out.create_sheet(title=safe_name)
        first = sess_rows[0]

        # Row1：A1=日期，C1=劇名（比照舊版）
        ws.row_dimensions[1].height = 22
        ws["A1"] = first["full_date"]
        ws["A1"].font = HDR_FONT
        ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
        ws["C1"] = show_name
        ws["C1"].font = HDR_FONT
        ws["C1"].alignment = Alignment(horizontal="left", vertical="center")

        # Row2：欄位標題
        ws.row_dimensions[2].height = 20
        col_headers = ["編號", "社群帳號", "姓名", "電話", "座位", "張數", "領取簽名"]
        for col, h in enumerate(col_headers, 1):
            c = ws.cell(row=2, column=col, value=h)
            c.font = SUB_FONT; c.fill = SUB_FILL
            c.alignment = CENTER; c.border = _thin_border()

        # 資料列
        for idx, r in enumerate(sess_rows):
            row_num = idx + 3
            fill = ALT_FILL if idx % 2 == 1 else None
            seat_str = "　".join(r["seats"])
            values = [r["id"], r["sns"], r["name"], r["tel"], seat_str, r["total"], ""]
            aligns = [CENTER, LEFT, CENTER, CENTER, LEFT, CENTER, CENTER]
            for col, (val, aln) in enumerate(zip(values, aligns), 1):
                c = ws.cell(row=row_num, column=col, value=val)
                c.font = BODY_FONT; c.alignment = aln; c.border = _thin_border()
                if fill: c.fill = fill

        # 合計列
        total_row = len(sess_rows) + 3
        ws.merge_cells(f"A{total_row}:E{total_row}")
        c = ws.cell(row=total_row, column=1, value=f"合計：{len(sess_rows)} 人")
        c.font = BOLD_FONT; c.alignment = CENTER; c.fill = TOT_FILL; c.border = _thin_border()
        tc = ws.cell(row=total_row, column=6, value=f"=SUM(F3:F{total_row-1})")
        tc.font = BOLD_FONT; tc.alignment = CENTER; tc.fill = TOT_FILL; tc.border = _thin_border()
        ws.cell(row=total_row, column=7).fill = TOT_FILL
        ws.cell(row=total_row, column=7).border = _thin_border()

        # 自動欄寬
        col_data = {
            1: [r["id"] for r in sess_rows] + ["編號"],
            2: [r["sns"] for r in sess_rows] + ["社群帳號"],
            3: [r["name"] for r in sess_rows] + ["姓名"],
            4: [r["tel"] for r in sess_rows] + ["電話"],
            5: ["　".join(r["seats"]) for r in sess_rows] + ["座位"],
            6: [r["total"] for r in sess_rows] + ["張數"],
            7: ["領取簽名"],
        }
        min_ws_map = {1:6, 2:12, 3:8, 4:14, 5:20, 6:6, 7:12}
        max_ws_map = {1:8, 2:25, 3:12, 4:16, 5:45, 6:8, 7:15}
        for col_idx, values in col_data.items():
            max_w = max((_cjk_len(v) for v in values if v), default=min_ws_map[col_idx])
            w = min(max(max_w + 2, min_ws_map[col_idx]), max_ws_map[col_idx])
            ws.column_dimensions[get_column_letter(col_idx)].width = w

        # 自動列高
        col5_w = ws.column_dimensions["E"].width
        for idx, r in enumerate(sess_rows):
            row_num = idx + 3
            seat_str = "　".join(r["seats"])
            lines = max(1, int(_cjk_len(seat_str) / max(col5_w, 1)) + 1)
            ws.row_dimensions[row_num].height = max(18, lines * 16)

        ws.freeze_panes = "A3"

    buf = BytesIO()
    wb_out.save(buf)
    buf.seek(0)
    return buf.getvalue()

def parse_session(raw: str):
    """解析日期場次字串，回傳 (排序key, 顯示文字)
    支援格式：
      2026/05/24(日)14:30、5/24 下午、2026-05-24 10:30
    """
    if not raw:
        return (99, 99, 99), ""
    raw = str(raw)
    # 優先嘗試 年/月/日 格式（避免誤把年份當月份）
    m = re.search(r'\d{4}[/\-](\d{1,2})[/\-](\d{1,2})', raw)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
    else:
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



# ══════════════════════════════════════════════════════════
# 解析函數：貴賓印製標籤版 ＆ 社福印製標籤版
#
# 欄位對應（已確認）：
#   Row0 = 劇名（略過）
#   Row1 = 標題列：編號、日期、姓名、座位（貴賓）/張數（社福）、張數
#   Row2 起 = 資料
#   A欄(0) = 編號
#   B欄(1) = 日期場次
#   C欄(2) = 姓名（已含「貴賓｜單位｜姓名」格式，直接使用）
#   D欄(3) = 座位（貴賓）或 張數（社福）
#   E欄(4) = 張數（貴賓）/ 無（社福）
#
# 標籤格式：{場次} {姓名欄原樣} X {張數}
# 唯一識別碼：{編號}_{工作表名稱}
# ══════════════════════════════════════════════════════════

def parse_label_sheet(df: pd.DataFrame, sheet_name: str, history_set: set):
    """
    解析貴賓印製標籤版 / 社福印製標籤版。
    Row0=劇名, Row1=標題, Row2起=資料
    """
    rows = df.values.tolist()
    results  = []
    warnings = []

    # 判斷是貴賓（有座位欄）還是社福（無座位欄）
    row1_str = " ".join(str(c) for c in rows[1]) if len(rows) > 1 else ""
    has_seat_col = "座位" in row1_str
    # 貴賓：A=編號 B=日期 C=姓名 D=座位 E=張數
    # 社福：A=編號 B=日期 C=姓名 D=張數
    COL_ID    = 0
    COL_DATE  = 1
    COL_NAME  = 2
    COL_COUNT = 4 if has_seat_col else 3

    for i, row in enumerate(rows):
        if i < 2: continue  # 跳過劇名列和標題列

        def get(idx):
            if idx >= len(row): return ""
            v = row[idx]
            if v is None: return ""
            if isinstance(v, __import__('datetime').datetime): return v.strftime("%Y/%m/%d")
            return str(v).strip()

        row_id   = re.sub(r"[^0-9]", "", get(COL_ID))
        name     = get(COL_NAME)
        date_raw = get(COL_DATE)
        count_raw = get(COL_COUNT)

        if not row_id or not name or not count_raw:
            continue

        count_clean = re.sub(r"[^0-9]", "", count_raw)
        if not count_clean: continue
        count = int(count_clean)
        if count <= 0: continue

        sort_key, display = parse_session(date_raw)
        key = f"{row_id}_{sheet_name}"
        label = f"{display} {name} X {count}"

        entry = {
            "key":              key,
            "id":               row_id,
            "name":             name,
            "sheet":            sheet_name,
            "total_count":      count,
            "count":            count,
            "earliest_sort":    sort_key,
            "earliest_display": display,
            "label":            label,
            "sns":              "",
            "tel":              "",
            "seats":            [],
            "is_new":           key not in history_set,
        }

        if entry["is_new"]:
            results.append(entry)
        else:
            entry_copy = dict(entry)
            results.append(entry_copy)

    # 分開 new / skipped
    tickets = [r for r in results if r["is_new"]]
    skipped = [r for r in results if not r["is_new"]]
    tickets.sort(key=lambda x: (x["earliest_sort"], int(x["id"])))
    skipped.sort(key=lambda x: (x["earliest_sort"], int(x["id"])))
    return tickets, skipped, warnings

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
            # C欄（社群帳號）若含「樓」「排」表示誤填了座位資訊，視為空白
            raw_sns = str(row[2]).strip() if len(row) > 2 and row[2] else ""
            sns_val = "" if any(x in raw_sns for x in ["樓", "排"]) else raw_sns
            merged[key] = {
                "key":              key,
                "id":               last_id,
                "name":             name or "(未填姓名)",
                "sns":              sns_val,
                "tel":              str(row[4]).strip() if len(row) > 4 and row[4] else "",
                "seats":            [],
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

        # 收集座位（同格內換行也分開）
        raw_seat = str(row[7]).strip() if len(row) > 7 and row[7] else ""
        if raw_seat:
            import re as _re
            for s in _re.split(r"[\n\r]+", raw_seat):
                s = s.strip()
                if s: merged[key]["seats"].append(s)
        merged[key]["total_count"] += count

    # 產生標籤
    tickets = []
    skipped = []
    for key, info in merged.items():
        label = (f"NO.{info['id']} {info['earliest_display']} "
                 f"貴賓｜{info['name']} X {info['total_count']}")
        entry = {**info, "label": label, "count": info["total_count"],
                  "sns": info.get("sns",""), "tel": info.get("tel",""),
                  "seats": info.get("seats",[])}
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
    "history_sid":    HISTORY_SHEET_ID,
    "history_set":    set(),
    "raw_sheets":     {},
    "selected_sheet": None,
    "rule_confirmed": False,
    "tickets":        [],
    "skipped":        [],
    "warnings":       [],
    "checked_keys":   set(),   # 使用者勾選「已列印」的 key 集合
    "uploaded_file_bytes": None,
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
    st.markdown("### ⚙️ 歷史紀錄")
    if not st.session_state.history_set:
        with st.spinner("載入歷史紀錄..."):
            st.session_state.history_set = load_history(HISTORY_SHEET_ID)

    if st.button("🔄 重新載入歷史", use_container_width=True):
        with st.spinner("連線中..."):
            st.session_state.history_set = load_history(HISTORY_SHEET_ID)
        st.success(f"已載入 {len(st.session_state.history_set)} 筆")

    st.divider()
    st.markdown(f"📦 歷史紀錄：**{len(st.session_state.history_set)}** 筆")
    st.divider()
    if st.button("🔄 重新開始", use_container_width=True):
        for k in ["raw_sheets","selected_sheet","rule_confirmed","tickets","skipped","warnings","checked_keys"]:
            st.session_state[k] = defaults[k]
        st.rerun()

    st.divider()
    st.markdown("### 🖨️ 補印")
    st.caption("輸入編號重新產生標籤（不受歷史紀錄限制）")
    reprint_id = st.text_input("編號", placeholder="例如：4 或 4,7,12", key="reprint_input")
    if st.button("產生補印標籤", use_container_width=True) and reprint_id:
        tickets_now = st.session_state.get("tickets", []) + st.session_state.get("skipped", [])
        ids = [x.strip() for x in reprint_id.replace("，",",").split(",") if x.strip()]
        found = [t for t in tickets_now if t["id"] in ids]
        if found:
            labels = "\n".join(t["label"] for t in found)
            st.code(labels, language=None)
        else:
            st.warning(f"找不到編號：{', '.join(ids)}\n（請先完成 STEP 2-3 載入資料）")


# ══════════════════════════════════════════════════════════
# STEP 1：上傳檔案
# ══════════════════════════════════════════════════════════
st.markdown('<div class="step-box">', unsafe_allow_html=True)
st.markdown('<div class="step-title">STEP 1 ｜ 上傳 xlsx 報名表</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("選取或拖曳 xlsx 檔案", type=["xlsx"], label_visibility="collapsed")
if uploaded:
    try:
        file_bytes = uploaded.read()
        st.session_state.uploaded_file_bytes = file_bytes
        xls = pd.ExcelFile(BytesIO(file_bytes))
        raw = {}
        for sname in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sname, header=None, dtype=str)
            raw[sname] = df.fillna("")
        if set(raw.keys()) != set(st.session_state.raw_sheets.keys()):
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
    row1_str = " ".join(str(c) for c in df.iloc[1].tolist()) if len(df) > 1 else ""

    # 判斷格式
    is_member  = ("姓名" in row0_str and "張數" in row0_str and "座位" in row0_str)
    is_label_vip  = ("編號" in row1_str and "日期" in row1_str and "張數" in row1_str and "座位" in row1_str)
    is_label_welfare = ("編號" in row1_str and "日期" in row1_str and "張數" in row1_str and "座位" not in row1_str and "姓名" not in row0_str)

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
                st.session_state.checked_keys   = set()
                st.rerun()

    elif is_label_vip or is_label_welfare:
        fmt_name = "貴賓印製標籤版" if is_label_vip else "社福印製標籤版"
        st.markdown(f'<div class="ok-box">✅ 偵測到：<strong>{fmt_name}</strong></div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            for key, val in [
                ("標題列位置", "第 1 行（Row 1）"),
                ("資料從",     "第 3 行開始"),
                ("列印條件",   "張數有值即列印"),
                ("合併邏輯",   "不合併，每行各自一筆"),
            ]:
                st.markdown(f'<div class="rule-box"><div class="rule-key">{key}</div><div class="rule-val">{val}</div></div>', unsafe_allow_html=True)
        with col2:
            for key, val in [
                ("唯一識別碼", "編號 ＋ 工作表名稱"),
                ("標籤格式",   "{日期場次} {姓名欄原樣} X {張數}"),
                ("排序方式",   "依日期場次排序"),
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
                    t, s, w = parse_label_sheet(df, sname, st.session_state.history_set)
                st.session_state.tickets        = t
                st.session_state.skipped        = s
                st.session_state.warnings       = w
                st.session_state.rule_confirmed = True
                st.session_state.checked_keys   = set()
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

        for session_display, sess_tickets in grouped.items():
            sess_people  = len(sess_tickets)
            sess_total   = sum(t["count"] for t in sess_tickets)
            sess_checked = sum(1 for t in sess_tickets if t["key"] in st.session_state.checked_keys)
            all_checked  = sess_checked == sess_people
            safe_key     = session_display.replace("/","_").replace(" ","_")
            unprinted    = [t for t in sess_tickets if t["key"] not in st.session_state.checked_keys]

            exp_label = (f"🗓 {session_display}　"
                         f"{sess_people} 人 ／ {sess_total} 張　"
                         f"{'✅ 全勾' if all_checked else f'已勾 {sess_checked}/{sess_people}'}")

            with st.expander(exp_label, expanded=False):
                # 全選按鈕 + 複製按鈕
                bcol1, bcol2 = st.columns([1, 1])
                with bcol1:
                    btn_label = "☐ 取消全選" if all_checked else "✅ 全選此場次"
                    if st.button(btn_label, key=f"sel_{safe_key}", use_container_width=True):
                        if not all_checked:
                            for t in sess_tickets:
                                st.session_state.checked_keys.add(t["key"])
                        else:
                            for t in sess_tickets:
                                st.session_state.checked_keys.discard(t["key"])
                        st.rerun()
                with bcol2:
                    if unprinted:
                        with st.popover(f"📋 複製未列印（{len(unprinted)} 筆）", use_container_width=True):
                            st.code("\n".join(t["label"] for t in unprinted), language=None)

                st.divider()

                # 每一筆勾選列
                for t in sess_tickets:
                    is_checked = t["key"] in st.session_state.checked_keys
                    col_chk, col_label = st.columns([1, 14])
                    with col_chk:
                        icon = "✅" if is_checked else "⬜"
                        if st.button(icon, key=f"chk_{t['key']}", use_container_width=True):
                            if is_checked:
                                st.session_state.checked_keys.discard(t["key"])
                            else:
                                st.session_state.checked_keys.add(t["key"])
                            st.rerun()
                    with col_label:
                        if is_checked:
                            st.markdown(
                                f'<p style="color:#bbb;font-size:0.88rem;font-family:monospace;margin:4px 0;">'                                f'<s>{t["label"]}</s></p>',
                                unsafe_allow_html=True
                            )
                        else:
                            st.markdown(
                                f'<p style="font-size:0.88rem;font-family:monospace;margin:4px 0;">{t["label"]}</p>',
                                unsafe_allow_html=True
                            )

        # 歸檔按鈕
        st.markdown("---")
        st.markdown("#### ✅ 列印完成後請點此歸檔")
        n_checked = len(st.session_state.checked_keys)
        col_p1, col_p2 = st.columns([3, 1])
        with col_p1:
            if n_checked == 0:
                st.info("請先在上方勾選已列印的票券，再點右側「歸檔」按鈕")
            else:
                st.info(f"已勾選 **{n_checked}** 筆，歸檔後下次不再出現。")
        with col_p2:
            st.markdown("<br>", unsafe_allow_html=True)
            if n_checked > 0:
                if st.button(f"✅ 歸檔勾選的 {n_checked} 筆", type="primary", use_container_width=True):
                    to_archive = [t for t in tickets if t["key"] in st.session_state.checked_keys]
                    keys   = [t["key"]   for t in to_archive]
                    labels = [t["label"] for t in to_archive]
                    if save_history(HISTORY_SHEET_ID, keys, labels):
                        st.session_state.history_set.update(keys)
                        st.session_state.tickets      = [t for t in tickets if t["key"] not in st.session_state.checked_keys]
                        st.session_state.checked_keys = set()
                        st.success(f"✅ 已歸檔 {len(keys)} 筆！")
                        st.rerun()

    if skipped:
        with st.expander(f"⏭️ 略過 {len(skipped)} 筆（歷史紀錄中已列印過）"):
            for t in skipped:
                st.markdown(f'<span style="color:#aaa;font-size:0.83rem;"><s>{t["label"]}</s></span>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# STEP 5：產生簽到表
# ══════════════════════════════════════════════════════════
if st.session_state.rule_confirmed and st.session_state.selected_sheet:
    st.markdown('<div class="step-box">', unsafe_allow_html=True)
    st.markdown('<div class="step-title">STEP 5 ｜ 產生簽到表（預覽 ＋ 下載 Excel）</div>', unsafe_allow_html=True)

    sname = st.session_state.selected_sheet
    row0_str = " ".join(str(c) for c in st.session_state.raw_sheets[sname].iloc[0].tolist())
    is_member = ("姓名" in row0_str and "張數" in row0_str and "座位" in row0_str)

    if is_member:
        # 演出名稱輸入 + 下載按鈕
        col_name, col_btn = st.columns([4, 1])
        with col_name:
            show_name_input = st.text_input(
                "演出名稱（顯示在簽到表標題）",
                value="親子音樂劇《阿甯咕的爸鼻不見了？》",
                key="show_name_input"
            )
        with col_btn:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.session_state.get("uploaded_file_bytes"):
                excel_bytes = generate_signin_excel(
                    st.session_state.uploaded_file_bytes,
                    sname,
                    show_name_input
                )
                st.download_button(
                    label="📥 下載 Excel",
                    data=excel_bytes,
                    file_name=f"簽到表_{sname}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary"
                )
            else:
                st.warning("請重新上傳 xlsx")

        st.markdown("---")

        # ── 網頁預覽：從已解析的 tickets+skipped 組合資料 ──
        all_tickets = st.session_state.tickets + st.session_state.skipped

        if all_tickets:
            # 重新依場次分組（含已列印的）
            preview_sessions = defaultdict(list)
            for t in sorted(all_tickets, key=lambda x: (x["earliest_sort"], int(x["id"]))):
                preview_sessions[t["earliest_display"]].append(t)

            st.markdown("#### 👁 簽到表預覽（所有場次）")
            st.caption("以下為網頁預覽版，實際 Excel 格式以下載為準")

            for session_display, sess_tickets in sorted(preview_sessions.items(), key=lambda x: x[1][0]["earliest_sort"]):
                sess_total = sum(t["count"] for t in sess_tickets)
                with st.expander(f"🗓 {session_display}　{len(sess_tickets)} 人 ／ {sess_total} 張", expanded=True):
                    # 表頭
                    st.markdown(
                        f'<div style="background:#1a1a2e;color:#e94560;font-weight:700;'
                        f'padding:0.4rem 0.75rem;border-radius:6px 6px 0 0;font-size:0.85rem;">'
                        f'{session_display}　{show_name_input}</div>',
                        unsafe_allow_html=True
                    )
                    # 表格資料
                    preview_data = []
                    for t in sess_tickets:
                        preview_data.append({
                            "編號":     t["id"],
                            "社群帳號": t.get("sns", ""),
                            "姓名":     t["name"],
                            "電話":     t.get("tel", ""),
                            "座位":     "　".join(t.get("seats", [])),
                            "張數":     t["count"],
                            "領取簽名": ""
                        })
                    df_preview = pd.DataFrame(preview_data)
                    st.dataframe(
                        df_preview,
                        use_container_width=True,
                        hide_index=True,
                        height=min(600, max(150, len(sess_tickets) * 36 + 40)),
                        column_config={
                            "編號":     st.column_config.TextColumn("編號",     width="small"),
                            "社群帳號": st.column_config.TextColumn("社群帳號", width="medium"),
                            "姓名":     st.column_config.TextColumn("姓名",     width="small"),
                            "電話":     st.column_config.TextColumn("電話",     width="medium"),
                            "座位":     st.column_config.TextColumn("座位",     width="large"),
                            "張數":     st.column_config.NumberColumn("張數",   width="small"),
                            "領取簽名": st.column_config.TextColumn("領取簽名", width="medium"),
                        }
                    )
                    st.caption(f"合計：{len(sess_tickets)} 人 ／ {sess_total} 張")
        else:
            st.info("請先完成 STEP 3 解析資料，才能預覽簽到表。")

    else:
        st.info("此工作表格式尚未支援產生簽到表，請聯絡管理員更新程式。")

    st.markdown('</div>', unsafe_allow_html=True)
