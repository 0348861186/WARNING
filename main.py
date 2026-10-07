import streamlit as st
import pandas as pd
import datetime
import io
import re
from openpyxl.styles import PatternFill

# Thư viện PDF
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

# Thư viện Gemini mới
try:
    from google import genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

st.set_page_config(
    page_title="Hệ Thống Đối Chiếu Xuất Kho & Cảnh Báo",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📦 Hệ Thống Đối Chiếu Lịch Xuất Hàng & Quản Lý Tồn Kho")
st.markdown("Tự động so sánh nhu cầu xuất kho, cảnh báo thiếu hàng trước hạn theo thuật toán phân bổ thực tế.")

# --- SIDEBAR ---
with st.sidebar:
    st.header("⚙️ Cấu hình & Tải File")
    gemini_api_key = st.text_input(
        "Google Gemini API Key (Tùy chọn):",
        type="password",
        help="Nhập key để kích hoạt phân tích thông minh từ Gemini."
    )
    
    st.divider()
    st.subheader("1. File Lịch Xuất Hàng")
    file_export = st.file_uploader("Chọn file lịch xuất (.xlsx, .csv)", type=["xlsx", "xls", "csv"], key="export")
    
    st.subheader("2. File Tồn Kho")
    file_stock = st.file_uploader("Chọn file tồn kho (.xlsx, .csv)", type=["xlsx", "xls", "csv"], key="stock")
    
    st.divider()
    alert_days = st.slider("Số ngày cảnh báo trước:", min_value=1, max_value=7, value=3)

# --- XỬ LÝ NGÀY THÁNG ---
def parse_custom_date(val):
    if pd.isna(val):
        return None
    val_str = str(val).strip()
    match = re.search(r'(\d{1,2})[/\-–\s]*thg\s*(\d{1,2})', val_str, re.IGNORECASE)
    current_year = datetime.datetime.now().year
    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        try:
            return datetime.date(current_year, month, day)
        except Exception:
            return None
    try:
        return pd.to_datetime(val_str).date()
    except Exception:
        return None

def read_file(file):
    if file.name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)

def strip_vietnamese_accent(text):
    """Hàm khử dấu tiếng Việt để xuất PDF không bị lỗi font ký tự"""
    import unicodedata
    nfkd = unicodedata.normalize('NFKD', str(text))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).replace('đ', 'd').replace('Đ', 'D')

# --- XUẤT EXCEL & PDF ---
def export_excel(df_detail, df_alerts):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_detail.to_excel(writer, sheet_name="DoiChieuChiTiet", index=False)
        df_alerts.to_excel(writer, sheet_name="CanhBaoThieuHang", index=False)
        
        wb = writer.book
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        # Highlight các dòng thiếu
        ws_alerts = wb["CanhBaoThieuHang"]
        for row in range(2, ws_alerts.max_row + 1):
            for col in range(1, ws_alerts.max_column + 1):
                ws_alerts.cell(row=row, column=col).fill = red_fill
                
    return output.getvalue()

def export_pdf(df_alerts):
    if not HAS_PDF:
        return None
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    elements = []
    
    title = Paragraph("<b>CANH BAO THIEU HANG XUAT KHO</b>", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 15))
    
    info_text = f"Ngay bao cao: {datetime.date.today().strftime('%d/%m/%Y')} | Don can xuat trong {alert_days} ngay toi:"
    elements.append(Paragraph(info_text, styles['Normal']))
    elements.append(Spacer(1, 15))
    
    pdf_data = [["So INV", "Ngay dong hang", "Ten san pham", "Can xuat", "Thieu"]]
    for _, row in df_alerts.iterrows():
        pdf_data.append([
            strip_vietnamese_accent(row.get("INV№", "")),
            str(row.get("箱詰めの日", "")),
            strip_vietnamese_accent(row.get("品名", ""))[:25],
            str(row.get("出荷数(CTNS)", "")),
            str(row.get("Số lượng thiếu", ""))
        ])
        
    table = Table(pdf_data, colWidths=[90, 95, 180, 65, 65])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D32F2F')),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BACKGROUND', (0,1), (-1,-1), colors.HexColor('#FFEBEE')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
    ]))
    elements.append(table)
    doc.build(elements)
    return buffer.getvalue()

# --- CHƯƠNG TRÌNH CHÍNH ---
if file_export and file_stock:
    try:
        df_exp_raw = read_file(file_export)
        df_stk_raw = read_file(file_stock)
        
        df_exp_raw.columns = [str(c).strip() for c in df_exp_raw.columns]
        df_stk_raw.columns = [str(c).strip() for c in df_stk_raw.columns]
        
        req_exp_cols = ["箱詰めの日", "INV№", "品名", "出荷数(CTNS)"]
        req_stk_cols = ["品名", "出荷数(CTNS)"]
        
        if not all(col in df_exp_raw.columns for col in req_exp_cols) or not all(col in df_stk_raw.columns for col in req_stk_cols):
            st.error("❌ Tên cột không khớp với yêu cầu mẫu!")
            st.stop()
            
        # 1. Tiền xử lý
        df_exp = df_exp_raw.copy()
        df_exp["箱詰めの日_raw"] = df_exp["箱詰めの日"].ffill()
        df_exp["INV№"] = df_exp["INV№"].ffill()
        df_exp = df_exp.dropna(subset=["品名"])
        df_exp["品名"] = df_exp["品名"].astype(str).str.strip()
        df_exp["出荷数(CTNS)"] = pd.to_numeric(df_exp["出荷数(CTNS)"], errors="coerce").fillna(0)
        df_exp["Parsed_Date"] = df_exp["箱詰めの日_raw"].apply(parse_custom_date)
        
        # Sắp xếp theo ngày xuất tăng dần để phân bổ theo FIFO
        df_exp = df_exp.sort_values(by="Parsed_Date", na_position='last').reset_index(drop=True)
        
        # 2. Xử lý tồn kho
        df_stk = df_stk_raw.copy().dropna(subset=["品名"])
        df_stk["品名"] = df_stk["品名"].astype(str).str.strip()
        df_stk["出荷数(CTNS)"] = pd.to_numeric(df_stk["出荷数(CTNS)"], errors="coerce").fillna(0)
        stock_dict = df_stk.groupby("品名")["出荷数(CTNS)"].sum().to_dict()
        
        # 3. Phân bổ tồn kho theo thứ tự ngày (FIFO logic)
        detail_records = []
        for _, row in df_exp.iterrows():
            item = row["品名"]
            req_qty = row["出荷数(CTNS)"]
            available = stock_dict.get(item, 0)
            
            if available >= req_qty:
                allocated = req_qty
                shortage = 0
                stock_dict[item] = available - req_qty
            else:
                allocated = max(0, available)
                shortage = req_qty - allocated
                stock_dict[item] = 0
                
            detail_records.append({
                "箱詰めの日": str(row["Parsed_Date"]) if row["Parsed_Date"] else str(row["箱詰めの日_raw"]),
                "INV№": row["INV№"],
                "品名": item,
                "出荷数(CTNS)": req_qty,
                "Tồn đáp ứng": allocated,
                "Số lượng thiếu": shortage,
                "Trạng thái": "Thiếu hàng" if shortage > 0 else "Đủ hàng",
                "_Parsed_Date": row["Parsed_Date"]
            })
            
        df_detail = pd.DataFrame(detail_records)
        
        # 4. Lọc cảnh báo gấp trước N ngày
        today = datetime.date.today()
        cutoff_date = today + datetime.timedelta(days=alert_days)
        
        df_alerts = df_detail[
            (df_detail["Số lượng thiếu"] > 0) & 
            (df_detail["_Parsed_Date"].notna()) & 
            (df_detail["_Parsed_Date"] <= cutoff_date)
        ].copy()
        
        # Hiển thị metrics
        m1, m2, m3 = st.columns(3)
        m1.metric("Tổng dòng xuất", len(df_detail))
        m2.metric("Số dòng bị thiếu", len(df_detail[df_detail["Số lượng thiếu"] > 0]), delta_color="inverse")
        m3.metric(f"🚨 Cảnh báo gấp (≤ {alert_days} ngày)", len(df_alerts), delta_color="inverse")
        
        # Hiển thị bảng cảnh báo
        st.subheader(f"🚨 Cảnh Báo Gấp (Hạn đóng hàng ≤ {cutoff_date.strftime('%d/%m/%Y')})")
        if not df_alerts.empty:
            st.error(f"Có **{len(df_alerts)}** đơn hàng sắp đến hạn xuất nhưng không đủ tồn kho:")
            
            show_cols = ["箱詰めの日", "INV№", "品名", "出荷数(CTNS)", "Tồn đáp ứng", "Số lượng thiếu"]
            st.dataframe(
                df_alerts[show_cols].style.apply(
                    lambda x: ['background-color: #ffcccc; color: #990000; font-weight: bold;'] * len(x), axis=1
                ),
                use_container_width=True
            )
        else:
            st.success(f"✅ An toàn! Toàn bộ đơn hàng trong {alert_days} ngày tới đều có đủ tồn kho.")
            
        # Tabs chi tiết & AI
        tab1, tab2 = st.tabs(["📋 Toàn bộ dữ liệu đối chiếu", "🤖 Phân tích cùng Gemini AI"])
        
        with tab1:
            st.dataframe(
                df_detail.drop(columns=["_Parsed_Date"]).style.apply(
                    lambda row: ['background-color: #ffebee'] * len(row) if row["Trạng thái"] == "Thiếu hàng" else [''] * len(row),
                    axis=1
                ),
                use_container_width=True
            )
            
        with tab2:
            if gemini_api_key and HAS_GEMINI:
                if st.button("Chạy phân tích rủi ro & giải pháp"):
                    with st.spinner("Gemini đang phân tích dữ liệu..."):
                        client = genai.Client(api_key=gemini_api_key)
                        prompt = f"""
                        Dưới đây là danh sách các đơn xuất khẩu bị thiếu hụt tồn kho trong {alert_days} ngày tới:
                        {df_alerts[["箱詰めの日", "INV№", "品名", "出荷数(CTNS)", "Số lượng thiếu"]].to_string(index=False)}

                        Hãy đưa ra:
                        1. Phân tích rủi ro theo số INV nào nguy cấp nhất.
                        2. 3 hành động cụ thể để xử lý (sắp xếp lại lịch, ưu tiên cont, hoặc liên hệ sản xuất).
                        Trả lời ngắn gọn, trực diện bằng tiếng Việt.
                        """
                        res = client.models.generate_content(model="gemini-2.5-flash", contents=prompt)
                        st.info(res.text)
            else:
                st.caption("Nhập Gemini API Key ở thanh bên trái để sử dụng tính năng này.")
                
        # Nút Download
        st.divider()
        c1, c2 = st.columns([1, 2])
        with c1:
            export_fmt = st.selectbox("Chọn định dạng xuất:", ["Excel (.xlsx)", "PDF (.pdf)"])
        with c2:
            st.write("")
            st.write("")
            if export_fmt == "Excel (.xlsx)":
                xl_data = export_excel(df_detail.drop(columns=["_Parsed_Date"]), df_alerts.drop(columns=["_Parsed_Date"]))
                st.download_button("📥 Tải báo cáo Excel", xl_data, f"Bao_Cao_{datetime.date.today()}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                pdf_data = export_pdf(df_alerts)
                if pdf_data:
                    st.download_button("📥 Tải báo cáo PDF", pdf_data, f"Canh_Bao_{datetime.date.today()}.pdf", "application/pdf")
                else:
                    st.error("Thiếu thư viện reportlab.")
                    
    except Exception as e:
        st.error(f"Đã xảy ra lỗi: {e}")
else:
    st.info("👈 Vui lòng tải lên cả 2 file ở thanh bên trái.")
