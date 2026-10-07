import ioimport refrom datetime import datetime, timedeltaimport pandas as pdimport streamlit as st# Tùy chọn Gemini APItry:

import google.generativeai as genai



HAS_GEMINI = Trueexcept ImportError:

HAS_GEMINI = False# Thư viện xuất PDFtry:

from reportlab.lib import colors

from reportlab.lib.pagesizes import letter

from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle



HAS_PDF = Trueexcept ImportError:

HAS_PDF = False# Cấu hình giao diện Streamlit

st.set_page_config(

page_title='Hệ Thống Đối Chiếu Xuất Kho & Cảnh Báo',

layout='wide',

initial_sidebar_state='expanded',

)



st.title('📦 Hệ Thống Đối Chiếu Lịch Xuất Hàng & Quản Lý Tồn Kho')

st.markdown(

'Tự động so sánh nhu cầu xuất kho, cảnh báo thiếu hàng trước 3 ngày và phân tích bằng AI.'

)# --- SIDEBAR: CẤU HÌNH & UPLOAD ---with st.sidebar:

st.header('⚙️ Cấu hình & Tải File')



# Nhập API Key Gemini (nếu muốn dùng)

gemini_api_key = st.text_input(

'Google Gemini API Key (Tùy chọn):',

type='password',

help='Nhập key để kích hoạt phân tích thông minh từ AI.',

)



st.divider()

st.subheader('1. File Lịch Xuất Hàng')

file_export = st.file_uploader(

'Chọn file lịch xuất (.xlsx, .csv)',

type=['xlsx', 'xls', 'csv'],

key='export',

)



st.subheader('2. File Tồn Kho')

file_stock = st.file_uploader(

'Chọn file tồn kho (.xlsx, .csv)',

type=['xlsx', 'xls', 'csv'],

key='stock',

)



st.divider()

alert_days = st.slider('Số ngày cảnh báo trước:', min_value=1, max_value=7, value=3)# --- HÀM HỖ TRỢ XỬ LÝ DỮ LIỆU ---def parse_vietnamese_date(val):

"""Hỗ trợ parse định dạng ngày dạng '4/thg 9' hoặc '22/thg 9' hoặc '2026-09-04'"""

if pd.isna(val):

return None

val_str = str(val).strip()



# Trường hợp dạng "4/thg 9" hoặc "17-thg 9"

match = re.search(r'(\d{1,2})[/\-–\s]*thg\s*(\d{1,2})', val_str, re.IGNORECASE)

current_year = datetime.now().year

if match:

day = int(match.group(1))

month = int(match.group(2))

try:

return datetime(current_year, month, day)

except Exception:

return None



# Thử parse chuẩn datetime

try:

return pd.to_datetime(val_str).to_pydatetime()

except Exception:

return Nonedef read_uploaded_file(file):

"""Đọc file Excel hoặc CSV"""

if file.name.endswith('.csv'):

return pd.read_csv(file)

else:

# Đọc header dòng 1 hoặc 2 tùy format, mặc định đọc dòng đầu tiên

return pd.read_excel(file)# --- HÀM TẠO FILE EXCEL VỚI ĐỊNH DẠNG TÔ ĐỎ ---def create_styled_excel(df_diff, df_alerts):

output = io.BytesIO()

with pd.ExcelWriter(output, engine='openpyxl') as writer:

# Sheet 1: Tổng hợp đối chiếu

df_diff.to_excel(writer, sheet_name='ChiTietDoiChieu', index=False)

# Sheet 2: Danh sách cảnh báo gấp

df_alerts.to_excel(writer, sheet_name='CanhBaoThieuHang', index=False)



# Tô màu đỏ các dòng thiếu trong openpyxl

workbook = writer.book

from openpyxl.styles import PatternFill



red_fill = PatternFill(

start_color='FFC7CE', end_color='FFC7CE', fill_type='solid'

)



ws_diff = workbook['ChiTietDoiChieu']

col_status_idx = None

for idx, col in enumerate(df_diff.columns, 1):

if col == 'Trạng thái':

col_status_idx = idx

break



if col_status_idx:

for row in range(2, len(df_diff) + 2):

if ws_diff.cell(row=row, column=col_status_idx).value == 'Thiếu hàng':

for col in range(1, len(df_diff.columns) + 1):

ws_diff.cell(row=row, column=col).fill = red_fill



return output.getvalue()# --- HÀM TẠO FILE PDF ---def create_pdf_report(df_alerts):

if not HAS_PDF:

return None

buffer = io.BytesIO()

doc = SimpleDocTemplate(buffer, pagesize=letter)

styles = getSampleStyleSheet()

elements = []



title = Paragraph(

'<b>BÁO CÁO CẢNH BÁO THIẾU HÀNG XUẤT KHO</b>', styles['Title']

)

elements.append(title)

elements.append(Spacer(1, 15))



intro_text = (

f'Thời gian lập: {datetime.now().strftime("%d/%m/%Y %H:%M")}<br/>Danh'

f' sách các mặt hàng thiếu hụt dự kiến xuất trong vòng {alert_days} ngày'

' tới:'

)

elements.append(Paragraph(intro_text, styles['Normal']))

elements.append(Spacer(1, 15))



# Dữ liệu bảng PDF

pdf_data = [[

'Số INV (INV№)',

'Tên hàng (品名)',

'Ngày đóng hàng',

'Số lượng thiếu',

]]

for _, row in df_alerts.iterrows():

pdf_data.append([

str(row.get('INV№', '')),

str(row.get('品名', ''))[:30],

str(row.get('箱詰めの日_hiển_thị', '')),

str(row.get('Số lượng thiếu', '')),

])



t = Table(pdf_data, colWidths=[110, 190, 110, 90])

t.setStyle(

TableStyle([

('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#D32F2F')),

('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),

('ALIGN', (0, 0), (-1, -1), 'CENTER'),

('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),

('BOTTOMPADDING', (0, 0), (-1, 0), 8),

('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#FFEBEE')),

('GRID', (0, 0), (-1, -1), 0.5, colors.grey),

])

)

elements.append(t)

doc.build(elements)

return buffer.getvalue()# --- CHƯƠNG TRÌNH CHÍNH ---if file_export is not None and file_stock is not None:

try:

df_exp_raw = read_uploaded_file(file_export)

df_stk_raw = read_uploaded_file(file_stock)



# Làm sạch tên cột

df_exp_raw.columns = [str(c).strip() for c in df_exp_raw.columns]

df_stk_raw.columns = [str(c).strip() for c in df_stk_raw.columns]



# Kiểm tra cột bắt buộc

req_exp_cols = ['箱詰めの日', 'INV№', '品名', '出荷数(CTNS)']

req_stk_cols = ['品名', '出荷数(CTNS)']



missing_exp = [c for c in req_exp_cols if c not in df_exp_raw.columns]

missing_stk = [c for c in req_stk_cols if c not in df_stk_raw.columns]



if missing_exp or missing_stk:

st.error(

f'❌ File không đúng định dạng cột yêu cầu:\n- File lịch xuất thiếu:'

f' {missing_exp}\n- File tồn kho thiếu: {missing_stk}'

)

st.stop()



# 1. Tiền xử lý File Lịch Xuất

df_exp = df_exp_raw.copy()

# Forward fill cho các dòng bị merge (như ngày và Inv№)

df_exp['箱詰めの日'] = df_exp['箱詰めの日'].ffill()

df_exp['INV№'] = df_exp['INV№'].ffill()



# Loại bỏ dòng không có tên hàng

df_exp = df_exp.dropna(subset=['品名'])

df_exp['品名'] = df_exp['品名'].astype(str).str.strip()

df_exp['出荷数(CTNS)'] = pd.to_numeric(

df_exp['出荷数(CTNS)'], errors='coerce'

).fillna(0)

df_exp['Parsed_Date'] = df_exp['箱詰めの日'].apply(parse_vietnamese_date)



# 2. Tiền xử lý File Tồn Kho

df_stk = df_stk_raw.copy().dropna(subset=['品名'])

df_stk['品名'] = df_stk['品名'].astype(str).str.strip()

df_stk['Tồn kho thực tế'] = pd.to_numeric(

df_stk['出荷数(CTNS)'], errors='coerce'

).fillna(0)



# Tổng hợp tồn kho theo mã hàng

df_stk_summary = (

df_stk.groupby('品名')['Tồn kho thực tế'].sum().reset_index()

)



# 3. Tính toán nhu cầu tổng theo từng mã hàng

df_exp_summary = (

df_exp.groupby('品名')['出荷数(CTNS)']

.sum()

.reset_index()

.rename(columns={'出荷数(CTNS)': 'Tổng xuất yêu cầu'})

)



# 4. Đối chiếu tồn kho so với tổng xuất

df_comparison = pd.merge(

df_exp_summary, df_stk_summary, on='品名', how='left'

)

df_comparison['Tồn kho thực tế'] = df_comparison[

'Tồn kho thực tế'

].fillna(0)

df_comparison['Chênh lệch (Tồn - Xuất)'] = (

df_comparison['Tồn kho thực tế'] - df_comparison['Tổng xuất yêu cầu']

)

df_comparison['Số lượng thiếu'] = df_comparison[

'Chênh lệch (Tồn - Xuất)'

].apply(lambda x: abs(x) if x < 0 else 0)

df_comparison['Trạng thái'] = df_comparison['Chênh lệch (Tồn - Xuất)'].apply(

lambda x: 'Thiếu hàng' if x < 0 else 'Đủ hàng'

)



# 5. Xác định cảnh báo trước N ngày theo INV№ và 品名

now = datetime.now()

cutoff_date = now + timedelta(days=alert_days)



# Ghép thông tin thiếu hàng ngược lại từng dòng xuất chi tiết

df_exp_detail = pd.merge(

df_exp,

df_comparison[['品名', 'Trạng thái', 'Chênh lệch (Tồn - Xuất)']],

on='品名',

how='left',

)



# Điều kiện cảnh báo: Bị thiếu hàng VÀ có ngày đóng hàng trong vòng N ngày tới (hoặc đã quá hạn)

urgent_alerts = df_exp_detail[

(df_exp_detail['Trạng thái'] == 'Thiếu hàng')

& (

df_exp_detail['Parsed_Date'].notna()

& (df_exp_detail['Parsed_Date'] <= cutoff_date)

)

].copy()



urgent_alerts['Số lượng thiếu'] = urgent_alerts[

'Chênh lệch (Tồn - Xuất)'

].apply(abs)

urgent_alerts['箱詰めの日_hiển_thị'] = urgent_alerts['箱詰めの日'].astype(

str

)



# --- HIỂN THỊ TRÊN STREAMLIT ---



# Metric Cards

col1, col2, col3 = st.columns(3)

total_items = len(df_comparison)

shortage_items = len(df_comparison[df_comparison['Trạng thái'] == 'Thiếu hàng'])

urgent_count = len(urgent_alerts['INV№'].unique())



col1.metric('Tổng số mã hàng', total_items)

col2.metric('Mã thiếu hàng', shortage_items, delta_color='inverse')

col3.metric(

f'Đơn INV cần đóng gấp ({alert_days} ngày)',

urgent_count,

delta_color='inverse',

)



# PHẦN CẢNH BÁO ĐỎ TRƯỚC 3 NGÀY

st.subheader(f'🚨 Cảnh Báo Gấp (Trước {alert_days} ngày đóng hàng)')

if not urgent_alerts.empty:

st.error(

f'Phát hiện **{len(urgent_alerts)}** mặt hàng thiếu hụt có lịch đóng'

f' trong vòng {alert_days} ngày tới!'

)



display_alert_df = urgent_alerts[[

'INV№',

'箱詰めの日',

'品名',

'出荷数(CTNS)',

'Số lượng thiếu',

]].rename(

columns={

'出荷数(CTNS)': 'Số lượng cần xuất đơn này',

'Số lượng thiếu': 'Tổng thiếu của mã này',

}

)



# Tô màu đỏ bảng cảnh báo

st.dataframe(

display_alert_df.style.apply(

lambda x: ['background-color: #ffcccc; color: #900;'] * len(x),

axis=1,

),

use_container_width=True,

)

else:

st.success(

f'✅ Không có đơn hàng nào bị thiếu hụt trong vòng {alert_days} ngày'

' tới.'

)



# PHẦN ĐỐI CHIẾU TỔNG THỂ

st.subheader('📊 Bảng Chi Tiết Đối Chiếu Tồn Kho')



def highlight_shortage(row):

if row['Trạng thái'] == 'Thiếu hàng':

return [

'background-color: #ffcdd2; color: #b71c1c; font-weight: bold;'

] * len(row)

return [''] * len(row)



styled_comp = df_comparison.style.apply(highlight_shortage, axis=1)

st.dataframe(styled_comp, use_container_width=True)



# PHẦN GEMINI AI PHÂN TÍCH

st.subheader('🤖 Đánh Giá & Khuyến Nghị Từ Gemini AI')

if gemini_api_key and HAS_GEMINI:

if st.button('Chạy phân tích AI'):

with st.spinner('Gemini AI đang phân tích rủi ro...'):

try:

genai.configure(api_key=gemini_api_key)

model = genai.GenerativeModel('gemini-1.5-flash')



# Chuẩn bị dữ liệu tóm tắt đưa cho AI

summary_prompt = f"""

Dưới đây là kết quả đối chiếu giữa lịch xuất hàng và tồn kho thực tế:

- Tổng số mặt hàng: {total_items}

- Số mặt hàng thiếu: {shortage_items}

- Danh sách các đơn xuất gấp bị thiếu trong {alert_days} ngày tới:

{urgent_alerts[['INV№', '品名', '出荷数(CTNS)', 'Số lượng thiếu']].to_string()}


Hãy đóng vai trò Chuyên gia Điều độ Sản xuất & Xuất nhập khẩu:

1. Đánh giá mức độ rủi ro đối với các Invoice sắp xuất.

2. Đưa ra 3 giải pháp xử lý cấp bách (chia cont, ưu tiên khách hàng, đẩy nhanh sản xuất,...).

Phản hồi ngắn gọn, súc tích bằng tiếng Việt.

"""

response = model.generate_content(summary_prompt)

st.info(response.text)

except Exception as e:

st.warning(f'Lỗi kết nối Gemini AI: {e}')

else:

st.caption(

'💡 *Gợi ý: Nhập Gemini API Key ở thanh bên trái để nhận đề xuất xử lý'

' thông minh.*'

)



# PHẦN DOWNLOAD BÁO CÁO (EXCEL / PDF)

st.divider()

st.subheader('📥 Tải Xuống Báo Cáo')

dl_col1, dl_col2 = st.columns([1, 2])



with dl_col1:

file_format = st.selectbox(

'Chọn định dạng xuất file:', ['Excel (.xlsx)', 'PDF (.pdf)']

)



with dl_col2:

st.write('')

st.write('')

if file_format == 'Excel (.xlsx)':

excel_bytes = create_styled_excel(df_comparison, urgent_alerts)

st.download_button(

label='Tải báo cáo Excel (Đã tô đỏ ô thiếu)',

data=excel_bytes,

file_name=f'BaoCao_ThieuHang_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx',

mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',

)

elif file_format == 'PDF (.pdf)':

if HAS_PDF:

pdf_bytes = create_pdf_report(urgent_alerts)

st.download_button(

label='Tải báo cáo cảnh báo PDF',

data=pdf_bytes,

file_name=f'CanhBao_XuatHang_{datetime.now().strftime("%Y%m%d_%H%M")}.pdf',

mime='application/pdf',

)

else:

st.error(

'Chưa cài đặt thư viện reportlab. Hãy cài đặt: `pip install'

' reportlab`'

)



except Exception as err:

st.error(f'Đã xảy ra lỗi trong quá trình xử lý: {err}')else:

st.info('👈 Vui lòng tải lên cả 2 file ở thanh bên trái để bắt đầu đối chiếu.') 

