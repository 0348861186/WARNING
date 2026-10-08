import streamlit as st
import pandas as pd
import datetime as dt
import io
import os
import re
import unicodedata
import urllib.request
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# SMART INVENTORY & SHIPMENT DASHBOARD (VIETNAMESE - CHINESE)
# 系统：出货排程对比与库存管理系统 (中越双语)
# ============================================================

APP_TITLE = "📦 Đối Chiếu Xuất Hàng & Tồn Kho | 出货排程对比与库存管理"
TODAY = dt.date.today()

# ------------------------------------------------------------
# OPTIONAL LIBRARIES
# ------------------------------------------------------------
try:
    from google import genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# ------------------------------------------------------------
# PAGE CONFIG
# ------------------------------------------------------------
st.set_page_config(
    page_title="Đối Chiếu Xuất Kho | 出货排程对比",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title(APP_TITLE)
st.caption(
    "Smart Import → Xử lý Merge Cell → FIFO Engine → Dashboard Song Ngữ (中越双语) → Xuất Excel & PDF"
)

# ------------------------------------------------------------
# COLUMN ALIASES
# ------------------------------------------------------------
COLUMN_ALIASES = {
    "date": [
        "新ETD", "装柜日期", "最初ETD", "ETD",
        "箱詰めの日", "箱詰め日", "Ngày đóng hàng", "Ngày xuất",
        "Ngày xuất hàng", "Ngày đóng", "Shipment Date", "Ship Date",
        "Packing Date", "Date", "Ngày", "日期", "出货日期"
    ],
    "inv": [
        "INV№", "INV", "Invoice", "Invoice No", "Invoice Number",
        "Số INV", "Số Invoice", "Mã INV", "INV No", "オーダー№", "发票号", "订单号"
    ],
    "item": [
        "品名", "品目", "Tên hàng", "Tên sản phẩm", "Sản phẩm",
        "Product Name", "Product", "Item", "Item Name", "SKU", "Mã hàng", "产品名称"
    ],
    "qty": [
        "注文数(CTNS)", "注文数（CTNS）", "注文数", "Số lượng đặt", "订单数",
        "出荷数(CTNS)", "出荷数（CTNS）", "出荷数", "出荷数量", "出货数",
        "Số lượng", "Số lượng xuất", "Số lượng (CTNS)",
        "Quantity", "Qty", "CTNS", "Shipment Qty", "Export Qty"
    ],
}

CANONICAL_NAMES = {
    "date": "箱詰めの日",
    "inv": "INV№",
    "item": "品名",
    "qty": "出荷数(CTNS)",
}

# ------------------------------------------------------------
# UI SIDEBAR
# ------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Cấu hình | 配置")

    gemini_api_key = st.text_input(
        "Google Gemini API Key (Tùy chọn | 可选)",
        type="password",
        help="Dùng Gemini để phân tích rủi ro tồn kho / 使用Gemini进行AI库存风险分析"
    )

    st.divider()

    st.subheader("1️⃣ Lịch xuất hàng | 出货计划")
    file_export = st.file_uploader(
        "Tải lên Excel / CSV (上传文件)",
        type=["xlsx", "xls", "csv"],
        key="export"
    )

    st.subheader("2️⃣ Tồn kho | 现存库存")
    file_stock = st.file_uploader(
        "Tải lên Excel / CSV (上传文件)",
        type=["xlsx", "xls", "csv"],
        key="stock"
    )

    st.divider()

    alert_days = st.slider(
        "🚨 Cảnh báo trước (ngày) | 提前预警天数",
        min_value=1,
        max_value=30,
        value=3
    )

    include_overdue = st.checkbox(
        "Bao gồm đơn đã quá hạn | 包括逾期订单",
        value=True
    )

    st.divider()
    st.caption("FIFO Inventory Engine • Song ngữ 中越双语")


# ============================================================
# TEXT / NORMALIZATION HELPERS
# ============================================================

def normalize_text(value):
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = text.replace(" ", " ")
    return text


def remove_accents(value):
    text = normalize_text(value)
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def clean_cell(value):
    if isinstance(value, pd.Series):
        value = value.iloc[0]
    if pd.isna(value):
        return None
    if isinstance(value, str):
        value = value.replace("\u00a0", " ").strip()
        if value == "":
            return None
    return value


# ============================================================
# FILE LOADER
# ============================================================

def get_excel_sheets(uploaded_file):
    if uploaded_file.name.lower().endswith(".csv"):
        return ["CSV"]

    try:
        uploaded_file.seek(0)
        xls = pd.ExcelFile(uploaded_file)
        sheets = xls.sheet_names
        uploaded_file.seek(0)
        return sheets
    except Exception:
        uploaded_file.seek(0)
        return []


def read_raw_file(uploaded_file, sheet_name=None):
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)

    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, header=None, dtype=object)
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, header=None, encoding="latin1", dtype=object)

    return pd.read_excel(
        uploaded_file,
        sheet_name=sheet_name or 0,
        header=None,
        dtype=object
    )


# ============================================================
# HEADER & COLUMN DETECTION
# ============================================================

def score_header_row(row_values, file_type):
    values = [normalize_text(v) for v in row_values if clean_cell(v) is not None]
    if not values:
        return 0

    score = 0
    for canonical, aliases in COLUMN_ALIASES.items():
        alias_norm = {normalize_text(x) for x in aliases}
        alias_ascii = {remove_accents(x) for x in aliases}

        for value in values:
            if value in alias_norm or remove_accents(value) in alias_ascii:
                score += 2
                break

    if len(values) >= 2:
        score += 1

    return score


def detect_header_row(raw_df, file_type="shipment"):
    max_scan = min(len(raw_df), 50)
    best_row = 0
    best_score = -1

    for idx in range(max_scan):
        score = score_header_row(raw_df.iloc[idx].tolist(), file_type)
        if score > best_score:
            best_score = score
            best_row = idx

    if best_score < 4:
        return None, best_score

    return best_row, best_score


def make_unique_columns(columns):
    result = []
    counter = {}

    for col in columns:
        base = str(col).strip() if clean_cell(col) is not None else "Unnamed"
        counter[base] = counter.get(base, 0) + 1

        if counter[base] == 1:
            result.append(base)
        else:
            result.append(f"{base}_{counter[base]}")

    return result


def find_column_by_alias(columns, aliases):
    normalized = {normalize_text(c): c for c in columns}
    ascii_map = {remove_accents(c): c for c in columns}

    for alias in aliases:
        a = normalize_text(alias)
        if a in normalized:
            return normalized[a]

    for alias in aliases:
        a = remove_accents(alias)
        if a in ascii_map:
            return ascii_map[a]

    for alias in aliases:
        a_norm = remove_accents(alias)
        if len(a_norm) >= 3:
            for col in columns:
                c_norm = remove_accents(col)
                if a_norm in c_norm:
                    return col

    return None


def detect_columns(df, file_type):
    columns = [str(c).strip() for c in df.columns]
    mapping = {}

    mapping["item"] = find_column_by_alias(columns, COLUMN_ALIASES["item"])
    mapping["qty"] = find_column_by_alias(columns, COLUMN_ALIASES["qty"])

    if file_type == "shipment":
        mapping["date"] = find_column_by_alias(columns, COLUMN_ALIASES["date"])
        mapping["inv"] = find_column_by_alias(columns, COLUMN_ALIASES["inv"])

    return mapping


def normalize_columns(df, mapping, file_type):
    df = df.copy()
    rename_map = {}

    for key, actual in mapping.items():
        if actual and actual in df.columns:
            rename_map[actual] = CANONICAL_NAMES[key]

    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated()]
    return df


# ============================================================
# PARSERS
# ============================================================

def parse_custom_date(value, reference_date=None):
    if reference_date is None:
        reference_date = TODAY

    if isinstance(value, pd.Series):
        value = value.dropna().iloc[0] if not value.dropna().empty else None

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        return None

    if isinstance(value, (pd.Timestamp, dt.datetime, dt.date)):
        return value.date() if hasattr(value, "date") else value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 100000:
            try:
                parsed = pd.to_datetime(value, unit="D", origin="1899-12-30")
                return parsed.date()
            except Exception:
                pass

    text = str(value).strip()
    if not text:
        return None

    month_en = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
    }
    match_en = re.search(r"(\d{1,2})[/\-\–\s]*([a-zA-Z]{3,})(?:[/\-\s]*(\d{4}))?", text)
    if match_en:
        day = int(match_en.group(1))
        m_str = match_en.group(2).lower()[:3]
        if m_str in month_en:
            month = month_en[m_str]
            year = int(match_en.group(3)) if match_en.group(3) else reference_date.year
            try:
                return dt.date(year, month, day)
            except ValueError:
                pass

    match = re.search(r"(\d{1,2})[/\-\–\s]*thg\.?\s*(\d{1,2})(?:[/\-\s]*(\d{4}))?", text, re.IGNORECASE)
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3)) if match.group(3) else reference_date.year
        try:
            return dt.date(year, month, day)
        except ValueError:
            return None

    iso_match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if iso_match:
        try:
            return dt.date(int(iso_match.group(1)), int(iso_match.group(2)), int(iso_match.group(3)))
        except ValueError:
            return None

    dmy_match = re.search(r"(\d{1,2})[./\-/](\d{1,2})[./\-/](\d{4})", text)
    if dmy_match:
        try:
            return dt.date(int(dmy_match.group(3)), int(dmy_match.group(2)), int(dmy_match.group(1)))
        except ValueError:
            return None

    try:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        pass

    return None


def parse_quantity(value):
    if isinstance(value, pd.Series):
        value = value.dropna().iloc[0] if not value.dropna().empty else None

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = re.sub(r"(?i)(ctns?|pcs?|cái|thùng|件|箱)\s*$", "", text).strip()
    text = text.replace("\u00a0", "")

    if "," in text and "." in text:
        if text.rfind(".") > text.rfind(","):
            text = text.replace(",", "")
        else:
            text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        parts = text.split(",")
        if len(parts[-1]) == 3 and all(p.isdigit() for p in parts):
            text = "".join(parts)
        else:
            text = text.replace(",", ".")
    elif "." in text:
        parts = text.split(".")
        if len(parts) > 2:
            text = "".join(parts)
        elif len(parts) == 2 and len(parts[-1]) == 3 and all(p.isdigit() for p in parts):
            text = "".join(parts)

    try:
        return float(text)
    except Exception:
        return None


# ============================================================
# PIPELINE IMPORT
# ============================================================

def clean_dataframe(df):
    df = df.copy()
    df = df.dropna(how="all")

    for col in df.columns:
        if isinstance(df[col], pd.DataFrame):
            df[col] = df[col].iloc[:, 0]
        df[col] = df[col].apply(clean_cell)

    return df.reset_index(drop=True)


def smart_import(uploaded_file, file_type, selected_sheet=None):
    raw = read_raw_file(uploaded_file, selected_sheet)
    if raw is None or raw.empty:
        return {"ok": False, "errors": ["File không có dữ liệu | 文件无数据。"]}

    header_row, header_score = detect_header_row(raw, file_type)
    if header_row is None:
        return {
            "ok": False,
            "errors": ["Không nhận diện được tiêu đề. Vui lòng kiểm tra cột bắt buộc | 无法识别标题行。"]
        }

    headers = make_unique_columns(raw.iloc[header_row].tolist())
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = clean_dataframe(df)

    mapping = detect_columns(df, file_type)
    required = ["item", "qty"]
    if file_type == "shipment":
        required += ["date", "inv"]

    missing = [x for x in required if not mapping.get(x)]
    if missing:
        return {"ok": False, "errors": [f"Thiếu các cột bắt buộc | 缺少必要列: {', '.join(missing)}"]}

    df = normalize_columns(df, mapping, file_type)
    warnings = []

    if file_type == "shipment":
        cols_to_ffill = [CANONICAL_NAMES["date"], CANONICAL_NAMES["inv"]]
        for col in cols_to_ffill:
            if col in df.columns:
                target_col = df[col]
                if isinstance(target_col, pd.Series):
                    before = target_col.isna().sum()
                    df[col] = target_col.ffill()
                    after = df[col].isna().sum()
                    if before > 0 and after < before:
                        warnings.append(f"Tự động điền {before - after} ô trống merge cell cho cột | 已自动填充合并单元格: '{col}'.")

    return {
        "ok": True,
        "errors": [],
        "warnings": warnings,
        "df": df,
        "mapping": mapping,
        "header_row": header_row,
    }


def validate_shipment(df):
    errors, warnings = [], []
    work = df.copy()

    for col in [CANONICAL_NAMES["item"], CANONICAL_NAMES["qty"], CANONICAL_NAMES["date"], CANONICAL_NAMES["inv"]]:
        if col in work.columns and isinstance(work[col], pd.DataFrame):
            work[col] = work[col].iloc[:, 0]

    empty_item = work["品名"].isna() | (work["品名"].astype(str).str.strip() == "")
    if empty_item.any():
        warnings.append(f"{int(empty_item.sum())} dòng không có tên sản phẩm (bỏ qua) | 缺少品名（已跳过）。")

    work = work[~empty_item].copy()
    work["品名"] = work["品名"].astype(str).str.strip()
    work["_Parsed_Date"] = work["箱詰めの日"].apply(parse_custom_date)
    work["出荷数(CTNS)"] = work["出荷数(CTNS)"].apply(parse_quantity)

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)
    if negative_qty.any():
        errors.append(f"{int(negative_qty.sum())} dòng có số lượng âm | 存在负数数量。")

    return work, errors, warnings


def validate_stock(df):
    errors, warnings = [], []
    work = df.copy()

    for col in [CANONICAL_NAMES["item"], CANONICAL_NAMES["qty"]]:
        if col in work.columns and isinstance(work[col], pd.DataFrame):
            work[col] = work[col].iloc[:, 0]

    empty_item = work["品名"].isna() | (work["品名"].astype(str).str.strip() == "")
    work = work[~empty_item].copy()
    work["品名"] = work["品名"].astype(str).str.strip()
    work["出荷数(CTNS)"] = work["出荷数(CTNS)"].apply(parse_quantity)

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)
    if negative_qty.any():
        errors.append(f"{int(negative_qty.sum())} dòng tồn kho có số lượng âm | 库存存在负数数量。")

    return work, errors, warnings


# ============================================================
# FIFO ENGINE & RISK CLASSIFICATION (BILINGUAL)
# ============================================================

def classify_risk_bilingual(shortage, date_value, today, alert_days):
    if pd.isna(shortage):
        return "⚫ Lỗi dữ liệu | 数据错误"
    if shortage <= 0:
        return "🟢 Đủ | 充足"
    if pd.isna(date_value) or date_value is None:
        return "⚫ Thiếu ngày | 缺少日期"

    days_left = (date_value - today).days
    if days_left < 0:
        return "🔴 Quá hạn | 逾期"
    if days_left == 0:
        return "🔴 Hôm nay | 今日出货"
    if days_left <= alert_days:
        return "🟠 Sắp đến hạn | 即将到期"

    return "🟡 Thiếu | 缺货"


def calculate_fifo(df_exp, df_stock, alert_days):
    work = df_exp.copy()
    work["_Original_Order"] = range(len(work))

    work = work.sort_values(
        by=["_Parsed_Date", "_Original_Order"],
        na_position="last"
    ).reset_index(drop=True)

    stock_dict = (
        df_stock.groupby("品名")["出荷数(CTNS)"]
        .sum()
        .fillna(0)
        .to_dict()
    )

    detail_records = []

    for _, row in work.iterrows():
        item = row["品名"]
        req_qty = row["出荷数(CTNS)"]

        if pd.isna(req_qty):
            continue

        available = float(stock_dict.get(item, 0))
        req_qty = float(req_qty)

        allocated = min(max(available, 0), req_qty)
        shortage = max(req_qty - allocated, 0)
        stock_dict[item] = max(available - req_qty, 0)

        if pd.notna(row["_Parsed_Date"]):
            date_display = row["_Parsed_Date"].strftime("%d/%b")
        else:
            date_display = str(row["箱詰めの日"])

        # Trạng thái song ngữ Việt - Trung
        status_label = "🟢 Đủ | 充足" if shortage == 0 else "🔴 Thiếu | 缺货"
        risk_label = classify_risk_bilingual(shortage, row["_Parsed_Date"], TODAY, alert_days)

        detail_records.append({
            "Ngày / 日期": date_display,
            "INV / 发票号": str(row["INV№"]),
            "Sản phẩm / 品名": item,
            "Cần xuất / 出货需求": int(req_qty) if req_qty.is_integer() else req_qty,
            "Tồn cấp / 库存分配": int(allocated) if allocated.is_integer() else allocated,
            "Thiếu / 缺货数": int(shortage) if shortage.is_integer() else shortage,
            "Trạng thái / 状态": status_label,
            "Rủi ro / 风险等级": risk_label,
            "_Parsed_Date": row["_Parsed_Date"],
        })

    return pd.DataFrame(detail_records), stock_dict


def build_alerts(df_detail, alert_days, include_overdue):
    df = df_detail.copy()
    base = (df["Thiếu / 缺货数"].fillna(0) > 0) & df["_Parsed_Date"].notna()
    future_or_today = (
        (df["_Parsed_Date"] >= TODAY) &
        (df["_Parsed_Date"] <= TODAY + dt.timedelta(days=alert_days))
    )
    overdue = df["_Parsed_Date"] < TODAY

    mask = base & (future_or_today | overdue) if include_overdue else base & future_or_today
    alerts = df[mask].copy()

    if not alerts.empty:
        alerts["Số ngày còn lại / 剩余天数"] = alerts["_Parsed_Date"].apply(lambda x: (x - TODAY).days)
        severity_order = {
            "🔴 Quá hạn | 逾期": 0, "🔴 Hôm nay | 今日出货": 1,
            "🟠 Sắp đến hạn | 即将到期": 2, "🟡 Thiếu | 缺货": 3
        }
        alerts["_RiskOrder"] = alerts["Rủi ro / 风险等级"].map(severity_order).fillna(9)
        alerts = alerts.sort_values(by=["_RiskOrder", "_Parsed_Date"]).drop(columns=["_RiskOrder"])

    return alerts


# ============================================================
# EXCEL EXPORT (UNICODE SAFE & SONG NGỮ)
# ============================================================

def autosize_worksheet(ws):
    for column_cells in ws.columns:
        max_length = 0
        col_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            try:
                # Tính độ dài an toàn cho cả ký tự CJK và Latin
                val = str(cell.value) if cell.value is not None else ""
                length = sum(2 if ord(c) > 127 else 1 for c in val)
                max_length = max(max_length, min(length, 60))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(14, max_length + 3)


def export_excel(df_detail, df_alerts, df_stock_summary, validation_warnings):
    output = io.BytesIO()

    detail_export = df_detail.drop(columns=["_Parsed_Date"], errors="ignore").copy()
    alert_export = df_alerts.drop(columns=["_Parsed_Date"], errors="ignore").copy()
    stock_export = df_stock_summary.copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_export.to_excel(writer, sheet_name="DoiChieu_对比明细", index=False)
        alert_export.to_excel(writer, sheet_name="CanhBao_预警明细", index=False)
        stock_export.to_excel(writer, sheet_name="TonKho_剩余库存", index=False)
        pd.DataFrame({
            "Cảnh báo dữ liệu / 数据检查警告": validation_warnings or ["Không có cảnh báo / 校验正常无警告。"]
        }).to_excel(writer, sheet_name="KiemTra_数据校验", index=False)

        wb = writer.book

        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        normal_font = Font(name="Segoe UI", size=10)

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        green_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")

        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.font = normal_font

            ws.freeze_panes = "A2"
            autosize_worksheet(ws)

        ws_detail = wb["DoiChieu_对比明细"]
        status_col = None
        for cell in ws_detail[1]:
            if "Trạng thái" in str(cell.value):
                status_col = cell.column

        if status_col:
            for row in range(2, ws_detail.max_row + 1):
                val = str(ws_detail.cell(row, status_col).value)
                if "Thiếu" in val or "缺货" in val:
                    for col in range(1, ws_detail.max_column + 1):
                        ws_detail.cell(row, col).fill = red_fill
                elif "Đủ" in val or "充足" in val:
                    ws_detail.cell(row, status_col).fill = green_fill

    return output.getvalue()


# ============================================================
# PDF EXPORT (FONT CJK & UNICODE ĐA NGÔN NGỮ)
# ============================================================

def setup_unicode_font():
    """
    Tự động dò và nạp font Unicode hỗ trợ cả tiếng Việt và chữ Hán (CJK).
    Nếu trên server/cloud không có, sẽ tự động tải font mã nguồn mở về thư mục làm việc.
    """
    system_fonts = [
        # Windows CJK & Unicode
        r"C:\Windows\Fonts\msyh.ttc",        # Microsoft YaHei
        r"C:\Windows\Fonts\msyh.ttf",
        r"C:\Windows\Fonts\simsun.ttc",      # SimSun
        r"C:\Windows\Fonts\simhei.ttf",      # SimHei
        r"C:\Windows\Fonts\arialuni.ttf",    # Arial Unicode MS
        r"C:\Windows\Fonts\seguisym.ttf",
        # Linux / Ubuntu CJK
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]

    for p in system_fonts:
        if Path(p).exists():
            try:
                pdfmetrics.registerFont(TTFont("BilingualFont", p))
                return "BilingualFont"
            except Exception:
                continue

    # Fallback: Tự động tải font Google Noto Sans SC về nếu đang chạy trên cloud/Linux container
    local_font_path = Path("NotoSansSC-Regular.ttf")
    if not local_font_path.exists():
        try:
            url = "https://github.com/google/fonts/raw/main/ofl/notosanssc/NotoSansSC%5Bwght%5D.ttf"
            urllib.request.urlretrieve(url, local_font_path)
        except Exception:
            pass

    if local_font_path.exists():
        try:
            pdfmetrics.registerFont(TTFont("BilingualFont", str(local_font_path)))
            return "BilingualFont"
        except Exception:
            pass

    return "Helvetica"


def export_pdf(df_detail):
    if not HAS_PDF:
        return None

    buffer = io.BytesIO()
    font_name = setup_unicode_font()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=15, leftMargin=15, topMargin=20, bottomMargin=20
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "PDFTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=15,
        leading=18
    )
    sub_style = ParagraphStyle(
        "PDFSub",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=8,
        leading=11
    )

    elements = [
        Paragraph("BÁO CÁO ĐỐI CHIẾU XUẤT HÀNG & TỒN KHO | 出货排程与库存对比报告", title_style),
        Spacer(1, 6),
        Paragraph(f"Ngày xuất / 报告日期: {TODAY.strftime('%d/%m/%Y')}", sub_style),
        Spacer(1, 10)
    ]

    headers = [
        "Ngày\n日期",
        "INV\n发票号",
        "Sản phẩm\n品名",
        "Cần xuất\n需求(CTN)",
        "Tồn cấp\n分配(CTN)",
        "Thiếu\n缺货(CTN)",
        "Trạng thái\n状态"
    ]
    data = [headers]

    for _, row in df_detail.iterrows():
        # Xóa emoji tròn để không xung đột với các font chữ Hán thuần túy
        status_clean = (
            str(row.get("Trạng thái / 状态", ""))
            .replace("🟢", "")
            .replace("🔴", "")
            .strip()
        )
        data.append([
            str(row.get("Ngày / 日期", "")),
            str(row.get("INV / 发票号", "")),
            str(row.get("Sản phẩm / 品名", ""))[:35],
            f"{row.get('Cần xuất / 出货需求', 0):,}",
            f"{row.get('Tồn cấp / 库存分配', 0):,}",
            f"{row.get('Thiếu / 缺货数', 0):,}",
            status_clean
        ])

    table = Table(data, repeatRows=1, colWidths=[65, 80, 240, 75, 75, 75, 85])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))

    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()


# ============================================================
# MAIN UI
# ============================================================

if not file_export or not file_stock:
    st.info("👈 Hãy tải lên cả file Lịch Xuất Hàng và Tồn Kho ở thanh bên trái | 请在左侧上传出货计划和库存文件。")
    st.markdown("""
### Tính năng hệ thống | 系统功能:
- ✅ Tự động gộp dữ liệu Merge Cell cho Ngày và Mã INV / 自动处理合并单元格
- ✅ Phân bổ tồn kho theo thuật toán FIFO ngày xuất / 遵循出货排程先进先出 (FIFO)
- ✅ Dashboard song ngữ trực quan: **Ngày / 日期 | INV / 发票号 | Sản phẩm / 品名 | Cần xuất / 出货需求 | Tồn cấp / 库存分配 | Thiếu / 缺货数 | Trạng thái / 状态**
- ✅ Xuất file Excel & PDF chuẩn font Unicode, không bị lỗi dấu tiếng Việt và chữ Hán / 导出无乱码 Excel 与 PDF
""")
    st.stop()

try:
    export_sheets = get_excel_sheets(file_export)
    stock_sheets = get_excel_sheets(file_stock)

    col1, col2 = st.columns(2)
    with col1:
        selected_export_sheet = st.selectbox(
            "📄 Sheet Lịch Xuất Hàng | 出货排程工作表",
            export_sheets,
            key="selected_export_sheet"
        )
    with col2:
        selected_stock_sheet = st.selectbox(
            "📄 Sheet Tồn Kho | 库存工作表",
            stock_sheets,
            key="selected_stock_sheet"
        )

    shipment_import = smart_import(
        file_export,
        "shipment",
        selected_export_sheet if selected_export_sheet != "CSV" else None
    )
    stock_import = smart_import(
        file_stock,
        "stock",
        selected_stock_sheet if selected_stock_sheet != "CSV" else None
    )

    if not shipment_import["ok"] or not stock_import["ok"]:
        st.error("❌ Không thể đọc file | 无法读取文件。")
        st.stop()

    df_exp, exp_errors, exp_warnings = validate_shipment(shipment_import["df"])
    df_stock, stock_errors, stock_warnings = validate_stock(stock_import["df"])

    all_errors = exp_errors + stock_errors
    all_warnings = (
        shipment_import.get("warnings", [])
        + stock_import.get("warnings", [])
        + exp_warnings
        + stock_warnings
    )

    if all_errors:
        for err in all_errors:
            st.error(err)
        st.stop()

    st.divider()

    if "calc_result" not in st.session_state:
        st.session_state["calc_result"] = None

    if st.button("🚀 XÁC NHẬN & CHẠY ĐỐI CHIẾU TỒN KHO | 确认并执行库存比对", type="primary", use_container_width=True):
        with st.spinner("⚙️ Đang phân bổ tồn kho FIFO | 正在进行先进先出库存分配..."):
            df_detail, remaining_stock = calculate_fifo(df_exp, df_stock, alert_days)
            df_alerts = build_alerts(df_detail, alert_days, include_overdue)
            stock_rows = [{"Sản phẩm / 品名": k, "Tồn kho còn lại / 剩余库存": v} for k, v in remaining_stock.items()]
            df_stock_summary = pd.DataFrame(stock_rows)

            st.session_state["calc_result"] = {
                "df_detail": df_detail,
                "df_alerts": df_alerts,
                "df_stock_summary": df_stock_summary,
                "all_warnings": all_warnings
            }

    if st.session_state["calc_result"] is None:
        st.info("Nhấn nút phía trên để bắt đầu đối chiếu | 点击上方按钮执行分析。")
        st.stop()

    res = st.session_state["calc_result"]
    df_detail = res["df_detail"]
    df_alerts = res["df_alerts"]
    df_stock_summary = res["df_stock_summary"]
    warnings_list = res["all_warnings"]

    # --------------------------------------------------------
    # METRICS (SONG NGỮ)
    # --------------------------------------------------------
    total_lines = len(df_detail)
    shortage_lines = int((df_detail["Thiếu / 缺货数"] > 0).sum())
    total_shortage = float(df_detail["Thiếu / 缺货数"].sum())

    m1, m2, m3 = st.columns(3)
    m1.metric("📋 Tổng dòng xuất | 出货总条目", f"{total_lines:,}")
    m2.metric("⚠️ Dòng thiếu | 缺货条目", f"{shortage_lines:,}")
    m3.metric("📦 Tổng thiếu | 缺货总量", f"{total_shortage:,.0f} CTNS")

    # --------------------------------------------------------
    # DASHBOARD BẢNG ĐỐI CHIẾU SONG NGỮ (7 CỘT CHUẨN)
    # --------------------------------------------------------
    st.subheader("📋 Bảng Đối Chiếu Lịch Xuất Hàng & Tồn Kho | 出货排程与库存对比表")

    display_cols = [
        "Ngày / 日期",
        "INV / 发票号",
        "Sản phẩm / 品名",
        "Cần xuất / 出货需求",
        "Tồn cấp / 库存分配",
        "Thiếu / 缺货数",
        "Trạng thái / 状态"
    ]
    df_dashboard = df_detail[display_cols].copy()

    st.dataframe(
        df_dashboard,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Ngày / 日期": st.column_config.TextColumn("Ngày / 日期", width="small"),
            "INV / 发票号": st.column_config.TextColumn("INV / 发票号", width="small"),
            "Sản phẩm / 品名": st.column_config.TextColumn("Sản phẩm / 品名", width="large"),
            "Cần xuất / 出货需求": st.column_config.NumberColumn("Cần xuất / 需求", format="%d"),
            "Tồn cấp / 库存分配": st.column_config.NumberColumn("Tồn cấp / 分配", format="%d"),
            "Thiếu / 缺货数": st.column_config.NumberColumn("Thiếu / 缺货", format="%d"),
            "Trạng thái / 状态": st.column_config.TextColumn("Trạng thái / 状态", width="medium"),
        }
    )

    # --------------------------------------------------------
    # NÚT XUẤT BÁO CÁO (EXCEL & PDF)
    # --------------------------------------------------------
    st.divider()
    st.subheader("📥 Xuất báo cáo | 导出报表")

    col_dl_a, col_dl_b = st.columns(2)

    with col_dl_a:
        excel_data = export_excel(df_detail, df_alerts, df_stock_summary, warnings_list)
        st.download_button(
            "📥 Tải Excel Báo Cáo | 下载 Excel 报表",
            data=excel_data,
            file_name=f"Bao_Cao_Xuat_Kho_{TODAY}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_dl_b:
        if HAS_PDF:
            pdf_data = export_pdf(df_detail)
            if pdf_data:
                st.download_button(
                    "📄 Tải PDF Báo Cáo | 下载 PDF 报表",
                    data=pdf_data,
                    file_name=f"Bao_Cao_Xuat_Kho_{TODAY}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.warning("Cần cài reportlab: pip install reportlab")

    # --------------------------------------------------------
    # CÁC TAB CHI TIẾT & GEMINI AI
    # --------------------------------------------------------
    st.divider()
    tab1, tab2 = st.tabs(["🏭 Tồn kho còn lại | 剩余库存", "🤖 Phân tích Gemini AI | 智能分析"])

    with tab1:
        st.dataframe(df_stock_summary, use_container_width=True, hide_index=True)

    with tab2:
        if not HAS_GEMINI:
            st.error("Chưa cài google-genai: pip install google-genai")
        elif not gemini_api_key:
            st.info("Nhập API Key ở thanh bên trái để sử dụng Gemini / 请在左侧输入API Key。")
        else:
            if st.button("🤖 Phân tích rủi ro đơn thiếu | 分析缺货风险", type="primary"):
                df_short = df_detail[df_detail["Thiếu / 缺货数"] > 0][display_cols]
                prompt = f"""
Bạn là chuyên gia điều phối kho và logistics xuất nhập khẩu.
Ngày hiện tại: {TODAY.strftime('%d/%m/%Y')}

Dưới đây là danh sách các đơn hàng xuất khẩu bị thiếu hàng (đã tính bằng FIFO):
{df_short.to_string(index=False)}

Hãy đưa ra đánh giá song ngữ (hoặc tiếng Việt có chú thích tiếng Trung cho các thuật ngữ chính):
1. 🔴 Đơn hàng nào có nguy cơ trễ tàu/chậm giao hàng cao nhất?
2. 📋 Đề xuất thứ tự ưu tiên sản xuất và đóng hàng.
3. 🛠️ 3 hành động cụ thể để xử lý số lượng thiếu ({total_shortage:,.0f} CTNS).
"""
                with st.spinner("Gemini đang phân tích | AI 正在分析中..."):
                    try:
                        client = genai.Client(api_key=gemini_api_key)
                        resp = client.models.generate_content(
                            model="gemini-2.5-flash",
                            contents=prompt
                        )
                        st.markdown(resp.text)
                    except Exception as e:
                        st.error(f"Lỗi: {e}")

except Exception as exc:
    st.error("❌ Đã xảy ra lỗi:")
    st.exception(exc)
