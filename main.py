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
# SMART INVENTORY & SHIPMENT DASHBOARD
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
    page_title="Đối Chiếu Xuất Kho",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title(APP_TITLE)
st.caption(
    "Smart Import → Tự sửa lỗi Merge Cell → Phân bổ FIFO → Dashboard Đối Chiếu → AI & Báo Cáo"
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
    st.caption("Auto FIFO Engine & Risk Detector")


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

    if isinstance(value, pd.Timestamp) or isinstance(value, dt.datetime) or isinstance(value, dt.date):
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
        return {"ok": False, "errors": ["File không có dữ liệu."]}

    header_row, header_score = detect_header_row(raw, file_type)
    if header_row is None:
        return {
            "ok": False,
            "errors": ["Không nhận diện được dòng tiêu đề. Vui lòng kiểm tra các cột bắt buộc."]
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
        return {"ok": False, "errors": [f"Thiếu các cột bắt buộc: {', '.join(missing)}"]}

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
                        warnings.append(f"Đã tự điền {before - after} ô trống merge cell cho cột '{col}'.")

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
        warnings.append(f"{int(empty_item.sum())} dòng không có tên sản phẩm (đã loại bỏ).")

    work = work[~empty_item].copy()
    work["品名"] = work["品名"].astype(str).str.strip()
    work["_Parsed_Date"] = work["箱詰めの日"].apply(parse_custom_date)
    work["出荷数(CTNS)"] = work["出荷数(CTNS)"].apply(parse_quantity)

    negative_qty = work["出荷数(CTNS)"].notna() & (work["出荷数(CTNS)"] < 0)
    if negative_qty.any():
        errors.append(f"{int(negative_qty.sum())} dòng có số lượng âm.")

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
        errors.append(f"{int(negative_qty.sum())} dòng tồn kho có số lượng âm.")

    return work, errors, warnings


# ============================================================
# FIFO ENGINE
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
            continue

        available = float(stock_dict.get(item, 0))
        req_qty = float(req_qty)

        allocated = min(max(available, 0), req_qty)
        shortage = max(req_qty - allocated, 0)
        stock_dict[item] = max(available - req_qty, 0)

        # Định dạng ngày: 01/Oct, 06/Oct...
        if pd.notna(row["_Parsed_Date"]):
            date_display = row["_Parsed_Date"].strftime("%d/%b")
        else:
            date_display = str(row["箱詰めの日"])

        detail_records.append({
            "Ngày": date_display,
            "INV": str(row["INV№"]),
            "Sản phẩm": item,
            "Cần xuất": int(req_qty) if req_qty.is_integer() else req_qty,
            "Tồn cấp": int(allocated) if allocated.is_integer() else allocated,
            "Thiếu": int(shortage) if shortage.is_integer() else shortage,
            "Trạng thái": "🟢 Đủ" if shortage == 0 else "🔴 Thiếu",
            "_Parsed_Date": row["_Parsed_Date"],
        })

    return pd.DataFrame(detail_records), stock_dict


# ============================================================
# MAIN UI
# ============================================================

if not file_export or not file_stock:
    st.info("👈 Hãy tải lên cả file Lịch Xuất Hàng và file Tồn Kho ở thanh bên trái.")
    st.markdown("""
### Tính năng hệ thống:
- ✅ Tự động gộp dữ liệu Merge Cell (INV, Ngày)
- ✅ Nhận diện chính xác cột `注文数(CTNS)` và `出荷数(CTNS)`
- ✅ Bảng đối chiếu trực quan chuẩn Dashboard: **Ngày | INV | Sản phẩm | Cần xuất | Tồn cấp | Thiếu | Trạng thái**
""")
    st.stop()

try:
    export_sheets = get_excel_sheets(file_export)
    stock_sheets = get_excel_sheets(file_stock)

    col1, col2 = st.columns(2)
    with col1:
        selected_export_sheet = st.selectbox("📄 Sheet Lịch Xuất Hàng", export_sheets, key="selected_export_sheet")
    with col2:
        selected_stock_sheet = st.selectbox("📄 Sheet Tồn Kho", stock_sheets, key="selected_stock_sheet")

    shipment_import = smart_import(file_export, "shipment", selected_export_sheet if selected_export_sheet != "CSV" else None)
    stock_import = smart_import(file_stock, "stock", selected_stock_sheet if selected_stock_sheet != "CSV" else None)

    if not shipment_import["ok"] or not stock_import["ok"]:
        st.error("❌ Không thể đọc file.")
        st.stop()

    df_exp, exp_errors, exp_warnings = validate_shipment(shipment_import["df"])
    df_stock, stock_errors, stock_warnings = validate_stock(stock_import["df"])

    all_errors = exp_errors + stock_errors
    if all_errors:
        for err in all_errors:
            st.error(err)
        st.stop()

    st.divider()

    if "calc_result" not in st.session_state:
        st.session_state["calc_result"] = None

    if st.button("🚀 XÁC NHẬN & CHẠY ĐỐI CHIẾU TỒN KHO", type="primary", use_container_width=True):
        with st.spinner("⚙️ Đang phân bổ tồn kho..."):
            df_detail, remaining_stock = calculate_fifo(df_exp, df_stock)
            stock_rows = [{"Sản phẩm": k, "Tồn kho còn lại": v} for k, v in remaining_stock.items()]
            df_stock_summary = pd.DataFrame(stock_rows)

            st.session_state["calc_result"] = {
                "df_detail": df_detail,
                "df_stock_summary": df_stock_summary,
            }

    if st.session_state["calc_result"] is None:
        st.info("Nhấn nút phía trên để bắt đầu đối chiếu dữ liệu.")
        st.stop()

    res = st.session_state["calc_result"]
    df_detail = res["df_detail"]
    df_stock_summary = res["df_stock_summary"]

    # --------------------------------------------------------
    # METRICS
    # --------------------------------------------------------
    total_lines = len(df_detail)
    shortage_lines = int((df_detail["Thiếu"] > 0).sum())
    total_shortage = float(df_detail["Thiếu"].sum())

    m1, m2, m3 = st.columns(3)
    m1.metric("📋 Tổng dòng xuất", f"{total_lines:,}")
    m2.metric("⚠️ Dòng thiếu hàng", f"{shortage_lines:,}")
    m3.metric("📦 Tổng thiếu", f"{total_shortage:,.0f} CTNS")

    # --------------------------------------------------------
    # DASHBOARD BẢNG ĐỐI CHIẾU
    # --------------------------------------------------------
    st.subheader("📋 Bảng Đối Chiếu Lịch Xuất Hàng & Tồn Kho")

    # Hiển thị chính xác 7 cột theo mẫu ảnh
    display_cols = ["Ngày", "INV", "Sản phẩm", "Cần xuất", "Tồn cấp", "Thiếu", "Trạng thái"]
    df_dashboard = df_detail[display_cols].copy()

    st.dataframe(
        df_dashboard,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Ngày": st.column_config.TextColumn("Ngày", width="small"),
            "INV": st.column_config.TextColumn("INV", width="small"),
            "Sản phẩm": st.column_config.TextColumn("Sản phẩm", width="large"),
            "Cần xuất": st.column_config.NumberColumn("Cần xuất", format="%d"),
            "Tồn cấp": st.column_config.NumberColumn("Tồn cấp", format="%d"),
            "Thiếu": st.column_config.NumberColumn("Thiếu", format="%d"),
            "Trạng thái": st.column_config.TextColumn("Trạng thái", width="small"),
        }
    )

    # --------------------------------------------------------
    # CÁC TAB CHI TIẾT & GEMINI AI
    # --------------------------------------------------------
    tab1, tab2 = st.tabs(["🏭 Tồn kho còn lại", "🤖 Phân tích Gemini AI"])

    with tab1:
        st.dataframe(df_stock_summary, use_container_width=True, hide_index=True)

    with tab2:
        if not HAS_GEMINI:
            st.error("Chưa cài google-genai: pip install google-genai")
        elif not gemini_api_key:
            st.info("Nhập API Key ở thanh bên trái để phân tích rủi ro.")
        else:
            if st.button("🤖 Phân tích rủi ro đơn thiếu bằng Gemini", type="primary"):
                df_short = df_detail[df_detail["Thiếu"] > 0][display_cols]
                prompt = f"""
Bạn là chuyên gia điều phối kho. Hãy đánh giá các đơn thiếu hàng sau:
{df_short.to_string(index=False)}

Đưa ra 3 đề xuất xử lý ngay:
1. Đơn nào nguy cấp nhất cần bổ sung sản xuất trước?
2. Phương án điều phối tồn kho thay thế.
3. Cảnh báo thời gian giao hàng.
"""
                with st.spinner("Gemini đang phân tích..."):
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
