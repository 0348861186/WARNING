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
# CODE 4 - SMART INVENTORY / SHIPMENT CHECK
# Nền tảng: Code 3
# Nâng cấp:
# - Smart Excel/CSV loader
# - Tự dò header
# - Tự nhận diện tên cột
# - Xử lý Merge Cell / ô trống bằng ffill
# - Chuẩn hóa ngày và số lượng
# - Validation trước khi tính
# - FIFO allocation theo ngày
# - Phân loại rủi ro
# - Gemini 2.5 Flash
# - Excel report
# - PDF report
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
# PAGE
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
# Chuẩn hóa nhiều cách đặt tên cột khác nhau về 4 trường chuẩn.
# ------------------------------------------------------------
COLUMN_ALIASES = {
    "date": [
        "箱詰めの日", "箱詰め日", "Ngày đóng hàng", "Ngày xuất",
        "Ngày xuất hàng", "Ngày đóng", "Shipment Date", "Ship Date",
        "Packing Date", "Date", "Ngày"
    ],
    "inv": [
        "INV№", "INV", "Invoice", "Invoice No", "Invoice Number",
        "Số INV", "Số Invoice", "Mã INV", "INV No"
    ],
    "item": [
        "品名", "品目", "Tên hàng", "Tên sản phẩm", "Sản phẩm",
        "Product Name", "Product", "Item", "Item Name", "SKU", "Mã hàng"
    ],
    "qty": [
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
    """Chuẩn hóa text để so sánh tên cột."""
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("　", " ")
    return text


def remove_accents(value):
    """Bỏ dấu để phục vụ fuzzy matching tên cột."""
    text = normalize_text(value)
    text = unicodedata.normalize("NFKD", text)
    return "".join(c for c in text if not unicodedata.combining(c))


def clean_cell(value):
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
    """Đọc danh sách sheet để người dùng biết file có nhiều sheet."""
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
    """Đọc file với header=None để tự dò header."""
    name = uploaded_file.name.lower()

    uploaded_file.seek(0)

    if name.endswith(".csv"):
        # thử UTF-8 trước, sau đó latin-1
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
    """Chấm điểm một dòng có khả năng là header."""
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

    # Ưu tiên dòng có nhiều text và ít số
    if len(values) >= 2:
        score += 1

    return score


def detect_header_row(raw_df, file_type="shipment"):
    """Tự tìm dòng header trong khoảng đầu file."""
    max_scan = min(len(raw_df), 50)

    best_row = 0
    best_score = -1

    for idx in range(max_scan):
        score = score_header_row(raw_df.iloc[idx].tolist(), file_type)
        if score > best_score:
            best_score = score
            best_row = idx

    # Ít nhất phải nhận diện được 2 trường
    if best_score < 5:
        return None, best_score

    return best_row, best_score


def make_unique_columns(columns):
    """Tránh lỗi khi Excel có tên cột trùng nhau."""
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

    # Fuzzy containment nhẹ
    for col in columns:
        c1 = remove_accents(col)
        for alias in aliases:
            c2 = remove_accents(alias)
            if len(c2) >= 4 and (c2 in c1 or c1 in c2):
                return col

    return None


def detect_columns(df, file_type):
    """Tự map cột thật → tên chuẩn."""
    columns = [str(c).strip() for c in df.columns]

    mapping = {}

    mapping["item"] = find_column_by_alias(columns, COLUMN_ALIASES["item"])
    mapping["qty"] = find_column_by_alias(columns, COLUMN_ALIASES["qty"])

    if file_type == "shipment":
        mapping["date"] = find_column_by_alias(columns, COLUMN_ALIASES["date"])
        mapping["inv"] = find_column_by_alias(columns, COLUMN_ALIASES["inv"])

    return mapping


def normalize_columns(df, mapping, file_type):
    """Đổi tên cột về schema chuẩn."""
    rename_map = {}

    for key, actual in mapping.items():
        if actual:
            rename_map[actual] = CANONICAL_NAMES[key]

    df = df.rename(columns=rename_map).copy()
    return df


# ============================================================
# SMART DATE PARSER
# ============================================================

def parse_custom_date(value, reference_date=None):
    """
    Hỗ trợ:
    4/thg 9
    04-thg 9
    04/09/2026
    2026-09-04
    datetime/date
    Excel serial date
    """
    if reference_date is None:
        reference_date = TODAY

    if value is None or pd.isna(value):
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
# SMART NUMBER PARSER
# ============================================================

def parse_quantity(value):
    """Chuẩn hóa số lượng từ Excel/CSV."""
    if value is None or pd.isna(value):
        return None

    if isinstance(value, bool):
        return None

    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    # loại đơn vị phổ biến
    text = re.sub(r"(?i)(ctns?|pcs?|cái|thùng)\s*$", "", text).strip()
    text = text.replace("\u00a0", "")

    # Xử lý:
    # 1,000 -> 1000
    # 1.000 -> 1000
    # 1,5 -> 1.5
    if "," in text and "." in text:
        # 1,234.56
        if text.rfind(".") > text.rfind(","):
            text = text.replace(",", "")
        else:
            # 1.234,56
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
        number = float(text)
        return number
    except Exception:
        return None


# ============================================================
# DATA NORMALIZATION
# ============================================================

def clean_dataframe(df):
    """Xóa dòng hoàn toàn rỗng và chuẩn hóa cell."""
    df = df.copy()
    df = df.dropna(how="all")

    for col in df.columns:
        df[col] = df[col].apply(clean_cell)

    return df.reset_index(drop=True)


def smart_import(uploaded_file, file_type, selected_sheet=None):
    """
    Pipeline:
    read raw → detect header → normalize columns → ffill → validate.
    """
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
                "Hãy kiểm tra file có chứa tên cột cần thiết hay không."
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
                "Thiếu: " + ", ".join(missing)
            ],
            "warnings": [],
            "df": df,
            "mapping": mapping,
            "header_row": header_row
        }

    df = normalize_columns(df, mapping, file_type)

    warnings = []

    # Merge Cell theo chiều dọc thường tạo NaN.
    # ffill chỉ áp dụng cho trường hợp phù hợp.
    if file_type == "shipment":
        for col in ["箱詰めの日", "INV№"]:
            if col in df.columns:
                before = df[col].isna().sum()
                df[col] = df[col].ffill()
                after = df[col].isna().sum()

                if before > 0 and after < before:
                    warnings.append(
                        f"Đã tự điền {before - after} ô trống của cột {col} (xử lý Merge Cell)."
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

    # Item
    empty_item = work["品名"].isna() | (work["品名"].astype(str).str.strip() == "")
    if empty_item.any():
        warnings.append(f"{int(empty_item.sum())} dòng không có tên sản phẩm; các dòng này sẽ bị bỏ.")

    work["品名"] = work["品名"].astype(str).str.strip()

    # Date
    work["_Parsed_Date"] = work["箱詰めの日"].apply(parse_custom_date)
    invalid_date = work["_Parsed_Date"].isna()

    if invalid_date.any():
        warnings.append(
            f"{int(invalid_date.sum())} dòng không đọc được ngày. "
            "Các dòng này vẫn giữ lại để kiểm tra nhưng không được đưa vào cảnh báo theo ngày."
        )

    # INV
    empty_inv = work["INV№"].isna() | (work["INV№"].astype(str).str.strip() == "")
    if empty_inv.any():
        warnings.append(
            f"{int(empty_inv.sum())} dòng không có INV. "
            "Cần kiểm tra vì có thể là dữ liệu bị thiếu."
        )

    # Quantity
    original_qty = work["出荷数(CTNS)"].copy()
    work["出荷数(CTNS)"] = original_qty.apply(parse_quantity)

    invalid_qty = work["出荷数(CTNS)"].isna()

    if invalid_qty.any():
        warnings.append(
            f"{int(invalid_qty.sum())} dòng có số lượng không hợp lệ hoặc trống. "
            "Các dòng này sẽ không được phân bổ."
        )

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)

    if negative_qty.any():
        errors.append(
            f"{int(negative_qty.sum())} dòng có số lượng âm. "
            "Hệ thống dừng để tránh tính sai tồn kho."
        )

    work = work[~empty_item].copy()

    return work, errors, warnings


def validate_stock(df):
    errors = []
    warnings = []
    work = df.copy()

    empty_item = work["品名"].isna() | (work["品名"].astype(str).str.strip() == "")
    if empty_item.any():
        warnings.append(f"{int(empty_item.sum())} dòng tồn kho không có tên sản phẩm.")

    work = work[~empty_item].copy()
    work["品名"] = work["品名"].astype(str).str.strip()

    work["出荷数(CTNS)"] = work["出荷数(CTNS)"].apply(parse_quantity)

    invalid_qty = work["出荷数(CTNS)"].isna()

    if invalid_qty.any():
        warnings.append(
            f"{int(invalid_qty.sum())} dòng tồn kho có số lượng không hợp lệ/trống."
        )

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)

    if negative_qty.any():
        errors.append(
            f"{int(negative_qty.sum())} dòng tồn kho có số lượng âm."
        )

    return work, errors, warnings


# ============================================================
# INVENTORY FIFO ENGINE
# ============================================================

def calculate_fifo(df_exp, df_stock):
    """
    FIFO theo ngày.
    Nếu cùng ngày thì giữ thứ tự xuất hiện trong file.
    """
    work = df_exp.copy()

    # Giữ thứ tự gốc để đảm bảo deterministic khi cùng ngày.
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

        # Không có số lượng hợp lệ
        if pd.isna(req_qty):
            detail_records.append({
                "箱詰めの日": row["箱詰めの日"],
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

    detail_export = df_detail.drop(
        columns=["_Parsed_Date"],
        errors="ignore"
    ).copy()

    alert_export = df_alerts.drop(
        columns=["_Parsed_Date"],
        errors="ignore"
    ).copy()

    stock_export = df_stock_summary.copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        detail_export.to_excel(
            writer,
            sheet_name="DoiChieuChiTiet",
            index=False
        )

        alert_export.to_excel(
            writer,
            sheet_name="CanhBao",
            index=False
        )

        stock_export.to_excel(
            writer,
            sheet_name="TonKhoConLai",
            index=False
        )

        pd.DataFrame({
            "CanhBao": validation_warnings or ["Không có cảnh báo dữ liệu."]
        }).to_excel(
            writer,
            sheet_name="KiemTraDuLieu",
            index=False
        )

        wb = writer.book

        # Header style
        header_fill = PatternFill(
            start_color="1F4E78",
            end_color="1F4E78",
            fill_type="solid"
        )

        red_fill = PatternFill(
            start_color="FFC7CE",
            end_color="FFC7CE",
            fill_type="solid"
        )

        orange_fill = PatternFill(
            start_color="FCE4D6",
            end_color="FCE4D6",
            fill_type="solid"
        )

        green_fill = PatternFill(
            start_color="E2F0D9",
            end_color="E2F0D9",
            fill_type="solid"
        )

        for ws in wb.worksheets:
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = Font(color="FFFFFF", bold=True)
                cell.alignment = Alignment(horizontal="center")

            ws.freeze_panes = "A2"
            autosize_worksheet(ws)

        # Highlight detail
        ws = wb["DoiChieuChiTiet"]

        status_col = None
        risk_col = None

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
            elif "Đủ" in str(status):
                if risk_col:
                    ws.cell(row, risk_col).fill = green_fill

            if "Quá hạn" in str(risk) or "Hôm nay" in str(risk):
                if risk_col:
                    ws.cell(row, risk_col).fill = red_fill
            elif "Sắp đến hạn" in str(risk):
                if risk_col:
                    ws.cell(row, risk_col).fill = orange_fill

        # Highlight alerts
        ws = wb["CanhBao"]
        for row in range(2, ws.max_row + 1):
            risk_value = ""

            for cell in ws[row]:
                if cell.value in [
                    "🔴 Quá hạn",
                    "🔴 Hôm nay",
                    "🟠 Sắp đến hạn",
                    "🟡 Thiếu hàng",
                ]:
                    risk_value = cell.value

            fill = (
                red_fill
                if "🔴" in str(risk_value)
                else orange_fill
                if "🟠" in str(risk_value)
                else red_fill
            )

            for col in range(1, ws.max_column + 1):
                ws.cell(row, col).fill = fill

    return output.getvalue()


# ============================================================
# PDF
# ============================================================

def find_unicode_font():
    candidates = [
        # Windows
        r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeui.ttf",
        r"C:\Windows\Fonts\NotoSans-Regular.ttf",

        # Linux
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
            pdfmetrics.registerFont(
                TTFont("SmartUnicode", font_path)
            )
            font_name = "SmartUnicode"
        except Exception:
            font_name = "Helvetica"

    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
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

    elements = []

    elements.append(
        Paragraph(
            "CẢNH BÁO THIẾU HÀNG XUẤT KHO",
            title_style
        )
    )

    elements.append(Spacer(1, 10))

    elements.append(
        Paragraph(
            f"Ngày báo cáo: {TODAY.strftime('%d/%m/%Y')} | "
            f"Cảnh báo trong {alert_days} ngày tới",
            normal_style
        )
    )

    elements.append(Spacer(1, 12))

    headers = [
        "INV",
        "Ngày đóng",
        "Sản phẩm",
        "Cần xuất",
        "Đáp ứng",
        "Thiếu",
        "Mức độ"
    ]

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

    table = Table(
        data,
        repeatRows=1,
        colWidths=[70, 75, 190, 65, 65, 65, 90]
    )

    table.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), font_name),
            ("FONTNAME", (0, 0), (-1, 0), font_name),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ])
    )

    elements.append(table)

    doc.build(elements)

    return buffer.getvalue()


# ============================================================
# GEMINI
# ============================================================

def run_gemini_analysis(api_key, df_alerts, df_detail, alert_days):
    if not HAS_GEMINI:
        return "Chưa cài thư viện google-genai."

    if not api_key:
        return "Chưa nhập Gemini API Key."

    alert_cols = [
        "箱詰めの日",
        "INV№",
        "品名",
        "出荷数(CTNS)",
        "Tồn đáp ứng",
        "Số lượng thiếu",
        "Mức độ",
    ]

    alert_text = (
        df_alerts[alert_cols]
        .to_string(index=False)
        if not df_alerts.empty
        else "Không có đơn thiếu trong vùng cảnh báo."
    )

    total_shortage = (
        df_detail["Số lượng thiếu"]
        .fillna(0)
        .sum()
    )

    prompt = f"""
Bạn là chuyên gia điều phối xuất hàng và tồn kho.

Ngày hiện tại: {TODAY.strftime('%d/%m/%Y')}

Dữ liệu dưới đây đã được Python tính toán bằng thuật toán FIFO.
KHÔNG được tự tính lại hoặc thay đổi số liệu.

Tổng số lượng thiếu: {total_shortage:,.0f} CTNS.

Danh sách cảnh báo:
{alert_text}

Hãy trả lời bằng tiếng Việt theo cấu trúc:

1. 🔴 ĐÁNH GIÁ NGUY CƠ
- INV nguy cấp nhất
- Lý do
- Số lượng thiếu
- Ngày đóng hàng

2. 📋 THỨ TỰ ƯU TIÊN XỬ LÝ
Xếp từ nguy cấp nhất đến ít nguy cấp hơn.

3. 🛠️ 3 HÀNH ĐỘNG ĐỀ XUẤT
- Ưu tiên sản xuất
- Kiểm tra tồn kho thực tế
- Điều chỉnh lịch xuất / liên hệ các bộ phận liên quan

4. ⚠️ LƯU Ý
Nếu dữ liệu có điểm bất thường, phải nói rõ.
Không được bịa thêm số liệu.
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
### Code 4 có thể tự xử lý:

- ✅ Header không nằm ở dòng đầu tiên
- ✅ Merge Cell / ô trống theo chiều dọc
- ✅ Nhiều cách đặt tên cột
- ✅ Ngày `4/thg 9`, `04/09/2026`, `2026-09-04`...
- ✅ Số lượng `1,000`, `1.000`, `100 CTNS`...
- ✅ Dòng trống
- ✅ Dữ liệu lỗi
- ✅ FIFO theo ngày
- ✅ Cảnh báo quá hạn / hôm nay / sắp đến hạn
""")

    st.stop()


# ============================================================
# IMPORT
# ============================================================

try:
    with st.spinner("🔎 Đang phân tích cấu trúc file..."):

        export_sheets = get_excel_sheets(file_export)
        stock_sheets = get_excel_sheets(file_stock)

    # Chọn sheet nếu Excel nhiều sheet
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

    # --------------------------------------------------------
    # IMPORT STATUS
    # --------------------------------------------------------
    st.divider()
    st.subheader("🔎 Kiểm tra file trước khi tính")

    c1, c2 = st.columns(2)

    with c1:
        if shipment_import["ok"]:
            st.success(
                f"✅ Lịch xuất: Header dòng {shipment_import['header_row'] + 1}"
            )

            if shipment_import.get("warnings"):
                for warning in shipment_import["warnings"]:
                    st.warning(warning)

            st.write("**Cột đã nhận diện:**")
            st.json(shipment_import["mapping"])

        else:
            st.error("❌ Không đọc được file lịch xuất.")
            for error in shipment_import["errors"]:
                st.error(error)

    with c2:
        if stock_import["ok"]:
            st.success(
                f"✅ Tồn kho: Header dòng {stock_import['header_row'] + 1}"
            )

            if stock_import.get("warnings"):
                for warning in stock_import["warnings"]:
                    st.warning(warning)

            st.write("**Cột đã nhận diện:**")
            st.json(stock_import["mapping"])

        else:
            st.error("❌ Không đọc được file tồn kho.")
            for error in stock_import["errors"]:
                st.error(error)

    if not shipment_import["ok"] or not stock_import["ok"]:
        st.stop()

    # --------------------------------------------------------
    # VALIDATE
    # --------------------------------------------------------
    df_exp, exp_errors, exp_warnings = validate_shipment(
        shipment_import["df"]
    )

    df_stock, stock_errors, stock_warnings = validate_stock(
        stock_import["df"]
    )

    all_errors = exp_errors + stock_errors
    all_warnings = (
        shipment_import.get("warnings", [])
        + stock_import.get("warnings", [])
        + exp_warnings
        + stock_warnings
    )

    if all_errors:
        st.error("🛑 Có lỗi dữ liệu nghiêm trọng. Hệ thống không chạy tính tồn kho.")

        for error in all_errors:
            st.error(error)

        st.info(
            "Hãy sửa dữ liệu nguồn rồi upload lại. "
            "Code 4 không tự đoán các giá trị có thể làm sai số liệu."
        )
        st.stop()

    if all_warnings:
        with st.expander(
            f"⚠️ Có {len(all_warnings)} cảnh báo dữ liệu - bấm để xem"
        ):
            for warning in all_warnings:
                st.warning(warning)

    # --------------------------------------------------------
    # PREVIEW DATA
    # --------------------------------------------------------
    with st.expander("👁️ Xem dữ liệu sau khi Code 4 chuẩn hóa", expanded=False):

        tab_a, tab_b = st.tabs([
            "📦 Lịch xuất",
            "🏭 Tồn kho"
        ])

        with tab_a:
            st.dataframe(
                df_exp.head(100),
                use_container_width=True
            )

        with tab_b:
            st.dataframe(
                df_stock.head(100),
                use_container_width=True
            )

    # --------------------------------------------------------
    # CALCULATE
    # --------------------------------------------------------
    st.divider()

    if st.button(
        "🚀 XÁC NHẬN & CHẠY ĐỐI CHIẾU TỒN KHO",
        type="primary",
        use_container_width=True
    ):
        st.session_state["run_analysis"] = True

    if not st.session_state.get("run_analysis", False):
        st.info("Kiểm tra dữ liệu phía trên, sau đó bấm nút xác nhận để chạy.")
        st.stop()

    with st.spinner("⚙️ Đang phân bổ tồn kho theo FIFO..."):
        df_detail, remaining_stock = calculate_fifo(
            df_exp,
            df_stock
        )

        df_detail = add_risk_level(
            df_detail,
            alert_days
        )

        df_alerts = build_alerts(
            df_detail,
            alert_days,
            include_overdue
        )

    # --------------------------------------------------------
    # STOCK SUMMARY
    # --------------------------------------------------------
    stock_rows = []

    for item, remaining in remaining_stock.items():
        stock_rows.append({
            "品名": item,
            "Tồn kho còn lại": remaining
        })

    df_stock_summary = pd.DataFrame(stock_rows)

    # --------------------------------------------------------
    # METRICS
    # --------------------------------------------------------
    total_lines = len(df_detail)
    shortage_lines = int(
        (df_detail["Số lượng thiếu"].fillna(0) > 0).sum()
    )
    total_shortage = float(
        df_detail["Số lượng thiếu"].fillna(0).sum()
    )
    alert_count = len(df_alerts)

    m1, m2, m3, m4 = st.columns(4)

    m1.metric(
        "📋 Tổng dòng xuất",
        f"{total_lines:,}"
    )

    m2.metric(
        "⚠️ Dòng thiếu",
        f"{shortage_lines:,}"
    )

    m3.metric(
        "📦 Tổng thiếu",
        f"{total_shortage:,.0f} CTNS"
    )

    m4.metric(
        "🚨 Cảnh báo",
        f"{alert_count:,}",
        delta_color="inverse"
    )

    # --------------------------------------------------------
    # RISK SUMMARY
    # --------------------------------------------------------
    st.subheader("🚦 Tổng quan mức độ rủi ro")

    risk_counts = (
        df_detail["Mức độ"]
        .value_counts()
        .rename_axis("Mức độ")
        .reset_index(name="Số dòng")
    )

    st.dataframe(
        risk_counts,
        use_container_width=True,
        hide_index=True
    )

    # --------------------------------------------------------
    # ALERTS
    # --------------------------------------------------------
    st.subheader(
        f"🚨 Cảnh báo thiếu hàng - {alert_days} ngày tới"
    )

    if df_alerts.empty:
        st.success(
            f"✅ Không có đơn thiếu hàng trong phạm vi cảnh báo {alert_days} ngày."
        )
    else:
        st.error(
            f"Phát hiện {len(df_alerts)} dòng cần xử lý."
        )

        alert_cols = [
            "箱詰めの日",
            "INV№",
            "品名",
            "出荷数(CTNS)",
            "Tồn đáp ứng",
            "Số lượng thiếu",
            "Số ngày còn lại",
            "Mức độ"
        ]

        st.dataframe(
            df_alerts[alert_cols],
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # FULL DETAIL
    # --------------------------------------------------------
    tab1, tab2, tab3 = st.tabs([
        "📋 Chi tiết đối chiếu",
        "🏭 Tồn kho còn lại",
        "🤖 Gemini AI"
    ])

    with tab1:
        display_detail = df_detail.drop(
            columns=["_Parsed_Date"],
            errors="ignore"
        )

        st.dataframe(
            display_detail,
            use_container_width=True,
            hide_index=True
        )

    with tab2:
        st.dataframe(
            df_stock_summary,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # GEMINI
    # --------------------------------------------------------
    with tab3:

        if not HAS_GEMINI:
            st.error(
                "Chưa cài google-genai. "
                "Cài bằng: pip install google-genai"
            )

        elif not gemini_api_key:
            st.info(
                "Nhập Gemini API Key ở thanh bên trái để dùng phân tích AI."
            )

        else:
            if st.button(
                "🤖 Phân tích rủi ro bằng Gemini",
                type="primary"
            ):
                with st.spinner("Gemini đang phân tích..."):
                    report = run_gemini_analysis(
                        gemini_api_key,
                        df_alerts,
                        df_detail,
                        alert_days
                    )

                st.markdown("### 🤖 Báo cáo AI")
                st.markdown(report)

                st.session_state["gemini_report"] = report

    # --------------------------------------------------------
    # EXPORT
    # --------------------------------------------------------
    st.divider()
    st.subheader("📥 Xuất báo cáo")

    col_a, col_b = st.columns(2)

    with col_a:
        excel_data = export_excel(
            df_detail,
            df_alerts,
            df_stock_summary,
            all_warnings
        )

        st.download_button(
            "📥 Tải Excel báo cáo",
            data=excel_data,
            file_name=f"Bao_Cao_Doi_Chieu_{TODAY}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    with col_b:
        if HAS_PDF:
            pdf_data = export_pdf(
                df_alerts,
                alert_days
            )

            if pdf_data:
                st.download_button(
                    "📄 Tải PDF cảnh báo",
                    data=pdf_data,
                    file_name=f"Canh_Bao_{TODAY}.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
        else:
            st.warning(
                "Chưa cài reportlab. Cài bằng: pip install reportlab"
            )

except Exception as exc:
    st.error("❌ Code 4 gặp lỗi khi xử lý file.")
    st.exception(exc)

    st.info(
        "Nếu lỗi xảy ra với một file Excel cụ thể, "
        "hãy kiểm tra phần traceback phía trên để xác định nguyên nhân."
    )
