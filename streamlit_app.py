import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl.styles import Border, Side, Alignment, Font

st.set_page_config(page_title="PDF to Master Summary", layout="wide")

st.title("🚀 Tool Chuyển Đổi PDF sang Bảng Tổng Hợp Master")
st.markdown("""
- **Trang 1:** Bỏ qua (Purchase Note).
- **Trang 2-29:** Tự động lấy SO, Ngày tháng và **Tên Store (Cột C)**.
- **Kết quả:** Một bảng dữ liệu sạch, có đóng khung để quản lý.
""")

uploaded_file = st.file_uploader("Tải file PDF 29 trang của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang phân tích cấu trúc ô và nhặt dữ liệu...'):
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            # Duyệt từ trang 2 (index 1) đến hết trang 29
            for i in range(1, len(pdf.pages)):
                page = pdf.pages[i]
                
                # 1. Lấy thông tin chung (SO, Order Date, Delivery Date) bằng văn bản
                text = page.extract_text()
                so_match = re.search(r"SO number:\s*(\d+)", text)
                order_date = re.search(r"Order date:\s*([\d/ :]+)", text)
                delivery_date = re.search(r"Delivery date:\s*([\d/]+)", text)
                
                so = so_match.group(1) if so_match else ""
                o_date = order_date.group(1).split()[0] if order_date else ""
                d_date = delivery_date.group(1) if delivery_date else ""

                # 2. LẤY TÊN STORE CHUẨN XÁC TẠI CỘT C
                store_name = ""
                # Trích xuất bảng đầu tiên (chứa địa chỉ)
                header_tables = page.extract_tables()
                if header_tables:
                    # Thông thường bảng Header là bảng đầu tiên của trang
                    header_data = header_tables[0]
                    # Dựa trên image_a4e7f5.png: 
                    # Dòng 3 (index 2), Cột C (index 2) là tên Store dưới chữ "For Store"
                    if len(header_data) >= 3 and len(header_data[2]) >= 3:
                        store_name = str(header_data[2][2]).replace('\n', ' ').strip()

                # 3. Trích xuất bảng chi tiết sản phẩm
                # Tìm tất cả các bảng và lọc bảng có chứa mã hàng
                for table in header_tables:
                    for row in table:
                        clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                        
                        # Điều kiện: Cột 1 là mã hàng (số) và độ dài >= 10
                        article_id = clean_row[0].strip()
                        if article_id.isdigit() and len(article_id) >= 10:
                            all_rows.append({
                                "SO Number": so,
                                "Order Date": o_date,
                                "Delivery Date": d_date,
                                "Store Name": store_name, # Cột D trong Excel kết quả
                                "Article": article_id,
                                "Description": clean_row[1],
                                "OU Qty": clean_row[5] if len(clean_row) > 5 else ""
                            })

        if all_rows:
            # Tạo DataFrame
            df_final = pd.DataFrame(all_rows)
            
            # Chuyển đổi số lượng sang kiểu số để tính toán nếu cần
            df_final["OU Qty"] = pd.to_numeric(df_final["OU Qty"], errors='coerce')
            
            # Ghi ra file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Master_Summary")
                
                # Định dạng bảng Excel (Đóng khung + Căn giữa Header)
                ws = writer.sheets["Master_Summary"]
                thin = Side(style='thin')
                border = Border(top=thin, left=thin, right=thin, bottom=thin)
                
                for row in ws.iter_rows(min_row=1, max_row=len(all_rows)+1, max_col=7):
                    for cell in row:
                        cell.border = border
                        if cell.row == 1:
                            cell.font = Font(bold=True)
                            cell.alignment = Alignment(horizontal="center")

            st.success(f"✅ Đã xử lý thành công {len(pdf.pages)} trang!")
            st.dataframe(df_final) # Hiển thị bản xem trước
            
            st.download_button(
                label="📥 Tải file Excel Tổng Hợp (Dữ liệu sạch)",
                data=output.getvalue(),
                file_name="Bao_cao_Tong_Hop_Master.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("⚠️ Không tìm thấy dữ liệu phù hợp trong file PDF.")
