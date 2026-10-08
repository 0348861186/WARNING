import streamlit as st
import pandas as pd
import datetime as dt
import io
import re
import unicodedata
from pathlib import Path
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# CODE 4 - SMART INVENTORY / SHIPMENT CHECK (BUGFIXED)
# Sửa lỗi:
# - ValueError: The truth value of a Series is ambiguous
# - Tránh trùng tên cột khi map canonical
# - Ép kiểu Series an toàn trong parse_quantity & parse_custom_date
# ============================================================

APP_TITLE = "📦 Hệ Thống Đối Chiếu Lịch Xuất Hàng & Quản Lý Tồn Kho"
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
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
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
    page_title="Đối Chiếu Xuất Kho",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title(APP_TITLE)
st.caption(
    "Smart Import → Kiểm tra dữ liệu → FIFO tồn kho → Cảnh báo rủi ro → Gemini AI → Báo cáo"
)

# ------------------------------------------------------------
# COLUMN ALIASES
# ------------------------------------------------------------
COLUMN_ALIASES = {
    "date": [
        "新ETD", "装柜日期", "最初ETD", "ETD",
        "箱詰めの日", "箱詰め日", "Ngày đóng hàng", "Ngày xuất",
        "Ngày xuất hàng", "Ngày đóng", "Shipment Date", "Ship Date",
        "Packing Date", "Date", "Ngày"
    ],
    "inv": [
        "INV№", "INV", "Invoice", "Invoice No", "Invoice Number",
        "Số INV", "Số Invoice", "Mã INV", "INV No", "オーダー№"
    ],
    "item": [
        "品名", "品目", "Tên hàng", "Tên sản phẩm", "Sản phẩm",
        "Product Name", "Product", "Item", "Item Name", "SKU", "Mã hàng"
    ],
    "qty": [
        "注文数(CTNS)", "注文数（CTNS）", "注文数", "Số lượng đặt",
        "出荷数(CTNS)", "出荷数（CTNS）", "出荷数", "出荷数量",
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
    st.header("⚙️ Cấu hình")

    gemini_api_key = st.text_input(
        "Google Gemini API Key (tùy chọn)",
        type="password",
        help="Dùng Gemini để phân tích rủi ro và đề xuất hành động."
    )

    st.divider()

    st.subheader("1️⃣ Lịch xuất hàng")
    file_export = st.file_uploader(
        "Upload Excel / CSV",
        type=["xlsx", "xls", "csv"],
        key="export"
    )

    st.subheader("2️⃣ Tồn kho")
    file_stock = st.file_uploader(
        "Upload Excel / CSV",
        type=["xlsx", "xls", "csv"],
        key="stock"
    )

    st.divider()

    alert_days = st.slider(
        "🚨 Cảnh báo trước (ngày)",
        min_value=1,
        max_value=30,
        value=3
    )

    include_overdue = st.checkbox(
        "Bao gồm đơn đã quá hạn",
        value=True
    )

    st.divider()
    st.caption("Code 4 - Smart Import / FIFO / Risk Engine")


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
# HEADER DETECTION
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


# ============================================================
# COLUMN MAPPING
# ============================================================

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

    # Dò cột theo thứ tự ưu tiên
    mapping["item"] = find_column_by_alias(columns, COLUMN_ALIASES["item"])
    mapping["qty"] = find_column_by_alias(columns, COLUMN_ALIASES["qty"])

    if file_type == "shipment":
        mapping["date"] = find_column_by_alias(columns, COLUMN_ALIASES["date"])
        mapping["inv"] = find_column_by_alias(columns, COLUMN_ALIASES["inv"])

    return mapping


def normalize_columns(df, mapping, file_type):
    """
    Chuyển đổi các cột đã tìm thấy về tên chuẩn và loại bỏ trùng lặp cột.
    """
    df = df.copy()
    rename_map = {}
    selected_cols = []

    for key, actual in mapping.items():
        if actual and actual in df.columns:
            canonical = CANONICAL_NAMES[key]
            rename_map[actual] = canonical
            selected_cols.append(actual)

    # Đổi tên và chỉ lấy đúng các cột chuẩn
    df = df.rename(columns=rename_map)
    canonical_list = list(rename_map.values())
    
    # Đảm bảo không bị trùng cột nếu có 2 cột cùng tên
    df = df.loc[:, ~df.columns.duplicated()]
    return df


# ============================================================
# SMART DATE PARSER (BUGFIX)
# ============================================================

def parse_custom_date(value, reference_date=None):
    if reference_date is None:
        reference_date = TODAY

    # Nếu truyền vào là Series do trùng cột, lấy phần tử đầu tiên
    if isinstance(value, pd.Series):
        value = value.dropna().iloc[0] if not value.dropna().empty else None

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except Exception:
        return None

    if isinstance(value, pd.Timestamp):
        return value.date()

    if isinstance(value, dt.datetime):
        return value.date()

    if isinstance(value, dt.date):
        return value

    # Excel serial date
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

    # Hỗ trợ: 1/Oct, 3-Oct, 16/Oct
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

    # 4/thg 9, 04-thg 9, 4 thg 9
    match = re.search(
        r"(\d{1,2})[/\-\–\s]*thg\.?\s*(\d{1,2})(?:[/\-\s]*(\d{4}))?",
        text,
        re.IGNORECASE
    )
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3)) if match.group(3) else reference_date.year
        try:
            return dt.date(year, month, day)
        except ValueError:
            return None

    # YYYY-MM-DD
    iso_match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
    if iso_match:
        try:
            return dt.date(
                int(iso_match.group(1)),
                int(iso_match.group(2)),
                int(iso_match.group(3))
            )
        except ValueError:
            return None

    # dd/mm/yyyy hoặc dd-mm-yyyy
    dmy_match = re.search(r"(\d{1,2})[./\-/](\d{1,2})[./\-/](\d{4})", text)
    if dmy_match:
        try:
            return dt.date(
                int(dmy_match.group(3)),
                int(dmy_match.group(2)),
                int(dmy_match.group(1))
            )
        except ValueError:
            return None

    try:
        parsed = pd.to_datetime(text, dayfirst=True, errors="coerce")
        if pd.notna(parsed):
            return parsed.date()
    except Exception:
        pass

    return None


# ============================================================
# SMART NUMBER PARSER (BUGFIX TRỰC TIẾP)
# ============================================================

def parse_quantity(value):
    # Nếu vô tình nhận phải Series, bóc giá trị đầu tiên
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

    text = re.sub(r"(?i)(ctns?|pcs?|cái|thùng)\s*$", "", text).strip()
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
# DATA NORMALIZATION
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
        return {
            "ok": False,
            "errors": ["File không có dữ liệu."],
            "warnings": [],
            "df": None,
            "mapping": {},
            "header_row": None
        }

    header_row, header_score = detect_header_row(raw, file_type)

    if header_row is None:
        return {
            "ok": False,
            "errors": [
                "Không thể tự nhận diện dòng tiêu đề.",
                "Hãy kiểm tra file có chứa các cột cần thiết hay không."
            ],
            "warnings": [],
            "df": None,
            "mapping": {},
            "header_row": None
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
        return {
            "ok": False,
            "errors": [
                "Không tìm thấy đủ các cột bắt buộc.",
                "Thiếu trường: " + ", ".join(missing)
            ],
            "warnings": [],
            "df": df,
            "mapping": mapping,
            "header_row": header_row
        }

    df = normalize_columns(df, mapping, file_type)
    warnings = []

    # Xử lý Merge Cell bằng ffill
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
                        warnings.append(
                            f"Đã tự điền {before - after} ô trống do merge cell cho cột '{col}'."
                        )

    return {
        "ok": True,
        "errors": [],
        "warnings": warnings,
        "df": df,
        "mapping": mapping,
        "header_row": header_row,
        "header_score": header_score
    }


# ============================================================
# VALIDATION
# ============================================================

def validate_shipment(df):
    errors = []
    warnings = []
    work = df.copy()

    # Đảm bảo các cột quan trọng không bị nhân bản (DataFrame con)
    for col in [CANONICAL_NAMES["item"], CANONICAL_NAMES["qty"], CANONICAL_NAMES["date"], CANONICAL_NAMES["inv"]]:
        if col in work.columns and isinstance(work[col], pd.DataFrame):
            work[col] = work[col].iloc[:, 0]

    # Item
    empty_item = work["品名"].isna() | (work["品名"].astype(str).str.strip() == "")
    if empty_item.any():
        warnings.append(f"{int(empty_item.sum())} dòng không có tên sản phẩm (đã loại bỏ).")

    work = work[~empty_item].copy()
    work["品名"] = work["品名"].astype(str).str.strip()

    # Date
    work["_Parsed_Date"] = work["箱詰めの日"].apply(parse_custom_date)
    invalid_date = work["_Parsed_Date"].isna()
    if invalid_date.any():
        warnings.append(
            f"{int(invalid_date.sum())} dòng không đọc được ngày xuất hàng."
        )

    # INV
    empty_inv = work["INV№"].isna() | (work["INV№"].astype(str).str.strip() == "")
    if empty_inv.any():
        warnings.append(f"{int(empty_inv.sum())} dòng bị thiếu mã INV.")

    # Qty
    work["出荷数(CTNS)"] = work["出荷数(CTNS)"].apply(parse_quantity)
    invalid_qty = work["出荷数(CTNS)"].isna()
    if invalid_qty.any():
        warnings.append(f"{int(invalid_qty.sum())} dòng số lượng bị trống hoặc sai định dạng.")

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)
    if negative_qty.any():
        errors.append(f"{int(negative_qty.sum())} dòng có số lượng âm. Hệ thống tạm dừng.")

    return work, errors, warnings


def validate_stock(df):
    errors = []
    warnings = []
    work = df.copy()

    for col in [CANONICAL_NAMES["item"], CANONICAL_NAMES["qty"]]:
        if col in work.columns and isinstance(work[col], pd.DataFrame):
            work[col] = work[col].iloc[:, 0]

    empty_item = work["品名"].isna() | (work["品名"].astype(str).str.strip() == "")
    if empty_item.any():
        warnings.append(f"{int(empty_item.sum())} dòng tồn kho không có tên sản phẩm.")

    work = work[~empty_item].copy()
    work["品名"] = work["品名"].astype(str).str.strip()

    work["出荷数(CTNS)"] = work["出荷数(CTNS)"].apply(parse_quantity)
    invalid_qty = work["出荷数(CTNS)"].isna()
    if invalid_qty.any():
        warnings.append(f"{int(invalid_qty.sum())} dòng tồn kho có số lượng không hợp lệ.")

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)
    if negative_qty.any():
        errors.append(f"{int(negative_qty.sum())} dòng tồn kho có số lượng âm.")

    return work, errors, warnings


# ============================================================
# INVENTORY FIFO ENGINE
# ============================================================

def calculate_fifo(df_exp, df_stock):
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
            detail_records.append({
                "箱詰めの日": str(row["箱詰めの日"]),
                "INV№": row["INV№"],
                "品名": item,
                "出荷数(CTNS)": None,
                "Tồn đáp ứng": 0,
                "Số lượng thiếu": None,
                "Trạng thái": "Dữ liệu lỗi",
                "Mức độ": "⚫ Dữ liệu lỗi",
                "_Parsed_Date": row["_Parsed_Date"],
            })
            continue

        available = float(stock_dict.get(item, 0))
        req_qty = float(req_qty)

        allocated = min(max(available, 0), req_qty)
        shortage = max(req_qty - allocated, 0)

        stock_dict[item] = max(available - req_qty, 0)

        detail_records.append({
            "箱詰めの日": (
                row["_Parsed_Date"].strftime("%d/%m/%Y")
                if pd.notna(row["_Parsed_Date"])
                else str(row["箱詰めの日"])
            ),
            "INV№": row["INV№"],
            "品名": item,
            "出荷数(CTNS)": req_qty,
            "Tồn đáp ứng": allocated,
            "Số lượng thiếu": shortage,
            "Trạng thái": "Thiếu hàng" if shortage > 0 else "Đủ hàng",
            "Mức độ": "",
            "_Parsed_Date": row["_Parsed_Date"],
        })

    return pd.DataFrame(detail_records), stock_dict


# ============================================================
# RISK ENGINE
# ============================================================

def classify_risk(row, today, alert_days):
    shortage = row.get("Số lượng thiếu", 0)
    date_value = row.get("_Parsed_Date")

    if pd.isna(shortage):
        return "⚫ Dữ liệu lỗi"

    if shortage <= 0:
        return "🟢 Đủ hàng"

    if pd.isna(date_value) or date_value is None:
        return "⚫ Thiếu ngày"

    days_left = (date_value - today).days

    if days_left < 0:
        return "🔴 Quá hạn"
    if days_left == 0:
        return "🔴 Hôm nay"
    if days_left <= alert_days:
        return "🟠 Sắp đến hạn"

    return "🟡 Thiếu hàng"


def add_risk_level(df_detail, alert_days):
    df = df_detail.copy()
    df["Mức độ"] = df.apply(
        lambda row: classify_risk(row, TODAY, alert_days),
        axis=1
    )
    return df


# ============================================================
# REPORT DATA
# ============================================================

def build_alerts(df_detail, alert_days, include_overdue):
    df = df_detail.copy()

    base = (
        (df["Số lượng thiếu"].fillna(0) > 0) &
        df["_Parsed_Date"].notna()
    )

    future_or_today = (
        (df["_Parsed_Date"] >= TODAY) &
        (df["_Parsed_Date"] <= TODAY + dt.timedelta(days=alert_days))
    )

    overdue = df["_Parsed_Date"] < TODAY

    if include_overdue:
        mask = base & (future_or_today | overdue)
    else:
        mask = base & future_or_today

    alerts = df[mask].copy()

    if not alerts.empty:
        alerts["Số ngày còn lại"] = alerts["_Parsed_Date"].apply(
            lambda x: (x - TODAY).days
        )

        severity_order = {
            "🔴 Quá hạn": 0,
            "🔴 Hôm nay": 1,
            "🟠 Sắp đến hạn": 2,
            "🟡 Thiếu hàng": 3,
        }

        alerts["_RiskOrder"] = alerts["Mức độ"].map(severity_order).fillna(9)
        alerts = alerts.sort_values(
            by=["_RiskOrder", "_Parsed_Date"]
        ).drop(columns=["_RiskOrder"])

    return alerts


# ============================================================
# EXCEL EXPORT
# ============================================================

def autosize_worksheet(ws):
    for column_cells in ws.columns:
        max_length = 0
        col_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells:
            try:
                length = len(str(cell.value)) if cell.value is not None else 0
                max_length = max(max_length, min(length, 60))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = max(10, max_length + 2)


def export_excel(df_detail, df_alerts, df_stock_summary, validation_warnings):
    output = io.BytesIO()

    detail_export = df_detail.drop(columns=["_Parsed_Date"], errors="ignore").copy()
    alert_export = df_alerts.drop(columns=["_Parsed_Date"], errors="ignore").copy()
    stock_export = df_stock_summary.copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_export.to_excel(writer, sheet_name="DoiChieuChiTiet", index=False)
        alert_export.to_excel(writer, sheet_name="CanhBao", index=False)
        stock_export.to_excel(writer, sheet_name="TonKhoConLai", index=False)
        pd.DataFrame({
            "CanhBao": validation_warnings or ["Không có cảnh báo dữ liệu."]
        }).to_excel(writer, sheet_name="KiemTraDuLieu", index=False)

        wb = writer.book

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        orange_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        green_fill = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")

        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")
            ws.freeze_panes = "A2"
            autosize_worksheet(ws)

        ws = wb["DoiChieuChiTiet"]
        status_col, risk_col = None, None
        for cell in ws[1]:
            if cell.value == "Trạng thái":
                status_col = cell.column
            if cell.value == "Mức độ":
                risk_col = cell.column

        for row in range(2, ws.max_row + 1):
            status = ws.cell(row, status_col).value if status_col else ""
            risk = ws.cell(row, risk_col).value if risk_col else ""

            if "Thiếu" in str(status):
                for col in range(1, ws.max_column + 1):
                    ws.cell(row, col).fill = red_fill
            elif "Đủ" in str(status) and risk_col:
                ws.cell(row, risk_col).fill = green_fill

            if ("Quá hạn" in str(risk) or "Hôm nay" in str(risk)) and risk_col:
                ws.cell(row, risk_col).fill = red_fill
            elif "Sắp đến hạn" in str(risk) and risk_col:
                ws.cell(row, risk_col).fill = orange_fill

    return output.getvalue()


# ============================================================
# PDF EXPORT
# ============================================================

def find_unicode_font():
    candidates = [
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\NotoSans-Regular.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/noto/NotoSans-Regular.ttf",
    ]
    for path in candidates:
        if Path(path).exists():
            return path
    return None


def export_pdf(df_alerts, alert_days):
    if not HAS_PDF:
        return None

    buffer = io.BytesIO()
    font_path = find_unicode_font()
    font_name = "Helvetica"

    if font_path:
        try:
            pdfmetrics.registerFont(TTFont("SmartUnicode", font_path))
            font_name = "SmartUnicode"
        except Exception:
            font_name = "Helvetica"

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "SmartTitle",
        parent=styles["Title"],
        fontName=font_name,
        fontSize=18,
        leading=22
    )
    normal_style = ParagraphStyle(
        "SmartNormal",
        parent=styles["Normal"],
        fontName=font_name,
        fontSize=9,
        leading=12
    )

    elements = [
        Paragraph("CẢNH BÁO THIẾU HÀNG XUẤT KHO", title_style),
        Spacer(1, 10),
        Paragraph(
            f"Ngày báo cáo: {TODAY.strftime('%d/%m/%Y')} | Cảnh báo trong {alert_days} ngày tới",
            normal_style
        ),
        Spacer(1, 12)
    ]

    headers = ["INV", "Ngày đóng", "Sản phẩm", "Cần xuất", "Đáp ứng", "Thiếu", "Mức độ"]
    data = [headers]

    for _, row in df_alerts.iterrows():
        data.append([
            str(row.get("INV№", "")),
            str(row.get("箱詰めの日", "")),
            str(row.get("品名", ""))[:35],
            f"{row.get('出荷数(CTNS)', 0):,.0f}",
            f"{row.get('Tồn đáp ứng', 0):,.0f}",
            f"{row.get('Số lượng thiếu', 0):,.0f}",
            str(row.get("Mức độ", ""))
        ])

    table = Table(data, repeatRows=1, colWidths=[70, 75, 190, 65, 65, 65, 90])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, -1), font_name),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))

    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()


# ============================================================
# GEMINI AI
# ============================================================

def run_gemini_analysis(api_key, df_alerts, df_detail, alert_days):
    if not HAS_GEMINI:
        return "Chưa cài thư viện google-genai. Vui lòng cài: pip install google-genai"

    if not api_key:
        return "Chưa nhập Gemini API Key."

    alert_cols = [
        "箱詰めの日", "INV№", "品名", "出荷数(CTNS)",
        "Tồn đáp ứng", "Số lượng thiếu", "Mức độ"
    ]

    alert_text = (
        df_alerts[alert_cols].to_string(index=False)
        if not df_alerts.empty
        else "Không có đơn thiếu trong vùng cảnh báo."
    )

    total_shortage = df_detail["Số lượng thiếu"].fillna(0).sum()

    prompt = f"""
Bạn là chuyên gia điều phối xuất hàng và tồn kho.
Ngày hiện tại: {TODAY.strftime('%d/%m/%Y')}

Dữ liệu dưới đây đã được phân bổ bằng thuật toán FIFO.
KHÔNG được tự tính lại hoặc thay đổi số liệu.

Tổng số lượng thiếu: {total_shortage:,.0f} CTNS.

Danh sách cảnh báo:
{alert_text}

Hãy trả lời bằng tiếng Việt theo cấu trúc sau:
1. 🔴 ĐÁNH GIÁ NGUY CƠ (INV nguy cấp nhất, lý do, số lượng thiếu, ngày đóng hàng)
2. 📋 THỨ TỰ ƯU TIÊN XỬ LÝ (Xếp từ nguy cấp nhất đến ít nguy cấp hơn)
3. 🛠️ 3 HÀNH ĐỘNG ĐỀ XUẤT (Sản xuất, kiểm tra thực tế, điều chỉnh lịch/báo khách)
4. ⚠️ LƯU Ý (Chỉ ra điểm bất thường nếu có, không bịa số liệu)
"""

    try:
        client = genai.Client(api_key=api_key)
        result = client.models.generate_content(
            model="gemini-2.5-flash",
            contents=prompt
        )
        return result.text if result and result.text else "Gemini không trả về nội dung."
    except Exception as exc:
        return f"Lỗi Gemini: {exc}"


# ============================================================
# MAIN UI
# ============================================================

if not file_export or not file_stock:
    st.info("👈 Hãy tải lên cả file Lịch Xuất Hàng và file Tồn Kho ở thanh bên trái.")
    st.markdown("""
### Hệ thống hỗ trợ xử lý thực tế:
- ✅ Tự động điền dữ liệu gộp ô (Merged Cells) cho Ngày và Mã INV
- ✅ Tương thích các cột ngày: `新ETD`, `装柜日期`, `最初ETD`, `箱詰めの日`...
- ✅ Tự nhận diện chính xác cột số lượng chi tiết `注文数(CTNS)`
- ✅ Phân bổ tồn kho theo thuật toán FIFO ngày xuất
- ✅ Báo cáo phân loại rủi ro: Quá hạn, Hôm nay, Sắp đến hạn
- ✅ Phân tích rủi ro chuyên sâu cùng Gemini 2.5 Flash
""")
    st.stop()


# ============================================================
# IMPORT PIPELINE
# ============================================================

try:
    export_sheets = get_excel_sheets(file_export)
    stock_sheets = get_excel_sheets(file_stock)

    col1, col2 = st.columns(2)
    with col1:
        selected_export_sheet = st.selectbox(
            "📄 Sheet Lịch Xuất Hàng",
            export_sheets,
            key="selected_export_sheet"
        )
    with col2:
        selected_stock_sheet = st.selectbox(
            "📄 Sheet Tồn Kho",
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

    st.divider()
    st.subheader("🔎 Kiểm tra cấu trúc file")

    c1, c2 = st.columns(2)
    with c1:
        if shipment_import["ok"]:
            st.success(f"✅ Lịch xuất: Header dòng {shipment_import['header_row'] + 1}")
            st.json(shipment_import["mapping"])
        else:
            st.error("❌ Lỗi file lịch xuất hàng:")
            for err in shipment_import["errors"]:
                st.error(err)

    with c2:
        if stock_import["ok"]:
            st.success(f"✅ Tồn kho: Header dòng {stock_import['header_row'] + 1}")
            st.json(stock_import["mapping"])
        else:
            st.error("❌ Lỗi file tồn kho:")
            for err in stock_import["errors"]:
                st.error(err)

    if not shipment_import["ok"] or not stock_import["ok"]:
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
        st.error("🛑 Dữ liệu có lỗi nghiêm trọng:")
        for err in all_errors:
            st.error(err)
        st.stop()

    if all_warnings:
        with st.expander(f"⚠️ Có {len(all_warnings)} cảnh báo cần lưu ý"):
            for w in all_warnings:
                st.warning(w)

    # --------------------------------------------------------
    # STATE & RUN ANALYSIS
    # --------------------------------------------------------
    st.divider()

    if "calc_result" not in st.session_state:
        st.session_state["calc_result"] = None

    if st.button("🚀 XÁC NHẬN & CHẠY ĐỐI CHIẾU TỒN KHO", type="primary", use_container_width=True):
        with st.spinner("⚙️ Đang phân bổ tồn kho theo FIFO..."):
            df_detail, remaining_stock = calculate_fifo(df_exp, df_stock)
            df_detail = add_risk_level(df_detail, alert_days)
            df_alerts = build_alerts(df_detail, alert_days, include_overdue)

            stock_rows = [{"品名": k, "Tồn kho còn lại": v} for k, v in remaining_stock.items()]
            df_stock_summary = pd.DataFrame(stock_rows)

            st.session_state["calc_result"] = {
                "df_detail": df_detail,
                "df_alerts": df_alerts,
                "df_stock_summary": df_stock_summary,
                "all_warnings": all_warnings
            }

    if st.session_state["calc_result"] is None:
        st.info("Hãy kiểm tra cấu trúc dữ liệu phía trên và nhấn nút để bắt đầu phân tích.")
        st.stop()

    res = st.session_state["calc_result"]
    df_detail = res["df_detail"]
    df_alerts = res["df_alerts"]
    df_stock_summary = res["df_stock_summary"]

    # --------------------------------------------------------
    # DASHBOARD METRICS
    # --------------------------------------------------------
    total_lines = len(df_detail)
    shortage_lines = int((df_detail["Số lượng thiếu"].fillna(0) > 0).sum())
    total_shortage = float(df_detail["Số lượng thiếu"].fillna(0).sum())
    alert_count = len(df_alerts)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📋 Tổng dòng xuất", f"{total_lines:,}")
    m2.metric("⚠️ Dòng thiếu", f"{shortage_lines:,}")
    m3.metric("📦 Tổng thiếu", f"{total_shortage:,.0f} CTNS")
    m4.metric("🚨 Cảnh báo", f"{alert_count:,}", delta_color="inverse")

    # --------------------------------------------------------
    # ALERTS DISPLAY
    # --------------------------------------------------------
    st.subheader(f"🚨 Cảnh báo thiếu hàng ({alert_days} ngày tới)")
    if df_alerts.empty:
        st.success(f"✅ Không có đơn thiếu hàng trong phạm vi cảnh báo {alert_days} ngày.")
    else:
        st.error(f"Phát hiện {len(df_alerts)} dòng cần xử lý gấp.")
        alert_cols = [
            "箱詰めの日", "INV№", "品名", "出荷数(CTNS)",
            "Tồn đáp ứng", "Số lượng thiếu", "Số ngày còn lại", "Mức độ"
        ]
        st.dataframe(df_alerts[alert_cols], use_container_width=True, hide_index=True)

    # --------------------------------------------------------
    # TABS: DETAIL / STOCK / GEMINI
    # --------------------------------------------------------
    tab1, tab2, tab3 = st.tabs(["📋 Chi tiết đối chiếu", "🏭 Tồn kho còn lại", "🤖 Gemini AI"])

    with tab1:
        st.dataframe(
            df_detail.drop(columns=["_Parsed_Date"], errors="ignore"),
            use_container_width=True,
            hide_index=True
        )

    with tab2:
        st.dataframe(df_stock_summary, use_container_width=True, hide_index=True)

    with tab3:
        if not HAS_GEMINI:
            st.error("Chưa cài google-genai. Cài đặt bằng: pip install google-genai")
        elif not gemini_api_key:
            st.info("Nhập Gemini API Key ở thanh bên trái để sử dụng tính năng phân tích AI.")
        else:
            if st.button("🤖 Phân tích rủi ro với Gemini", type="primary"):
                with st.spinner("Gemini đang phân tích dữ liệu..."):
                    report = run_gemini_analysis(gemini_api_key, df_alerts, df_detail, alert_days)
                    st.session_state["gemini_report"] = report

            if "gemini_report" in st.session_state:
                st.markdown("### 🤖 Báo cáo phân tích AI")
                st.markdown(st.session_state["gemini_report"])

    # --------------------------------------------------------
    # EXPORT REPORT
    # --------------------------------------------------------
    st.divider()
    st.subheader("📥 Xuất báo cáo")

    col_a, col_b = st.columns(2)
    with col_a:
        excel_data = export_excel(df_detail, df_alerts, df_stock_summary, all_warnings)
        st.download_button(
            "📥 Tải Excel báo cáo",
            data=excel_data,
            file_name=f"Bao_Cao_Doi_Chieu_{TODAY}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_b:
        if HAS_PDF:
            pdf_data = export_pdf(df_alerts, alert_days)
            if pdf_data:
                st.download_button(
                    "📄 Tải PDF cảnh báo",
                    data=pdf_data,
                    file_name=f"Canh_Bao_{TODAY}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.warning("Cần cài reportlab để xuất PDF: pip install reportlab")

except Exception as exc:
    st.error("❌ Đã xảy ra lỗi khi xử lý:")
    st.exception(exc)
