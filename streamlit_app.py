import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl.styles import Border, Side

st.set_page_config(page_title="PDF Master Summary", layout="wide")
st.title("📊 Master Tool: Chuyển đổi PDF sang Bảng Tổng Hợp")

uploaded_file = st.file_uploader("Tải file PDF 29 trang lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang nhặt dữ liệu và xử lý sạch...'):
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            # Duyệt từ trang 2 đến hết (Trang 1 là Purchase Note nên bỏ qua trong bản tổng hợp này)
            for i in range(1, len(pdf.pages)):
                page = pdf.pages[i]
                text = page.extract_text()
                
                # 1. Tìm thông tin chung của trang (SO, Date, Store) bằng Regex
                so_match = re.search(r"SO number:\s*(\d+)", text)
                order_date = re.search(r"Order date:\s*([\d/ :]+)", text)
                delivery_date = re.search(r"Delivery date:\s*([\d/]+)", text)
                # Lấy tên Store sau chữ "For Store"
                store_match = re.search(r"For Store\s*\n(.*?)\n", text)
                
                so = so_match.group(1) if so_match else ""
                o_date = order_date.group(1).split()[0] if order_date else "" # Chỉ lấy ngày
                d_date = delivery_date.group(1) if delivery_date else ""
                store = store_match.group(1).strip() if store_match else ""

                # 2. Trích xuất bảng sản phẩm
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # Làm sạch dữ liệu ô
                        clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                        
                        # Chỉ lấy dòng nếu cột đầu tiên là mã hàng (số)
                        article_id = clean_row[0].strip()
                        if article_id.isdigit() and len(article_id) >= 10:
                            # Tạo dòng dữ liệu mới theo đúng cấu trúc bạn muốn
                            new_row = {
                                "SO Number": so,
                                "Order Date": o_date,
                                "Delivery Date": d_date,
                                "Store": store,
                                "Article": article_id,
                                "Description": clean_row[1],
                                "OU Qty": clean_row[5] if len(clean_row) > 5 else ""
                            }
                            all_rows.append(new_row)

        if all_rows:
            df_final = pd.DataFrame(all_rows)
            
            # Xuất file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Master Summary")
                
                # Đóng khung cho đẹp
                ws = writer.sheets["Master Summary"]
                thin = Side(style='thin')
                for row in ws.iter_rows(min_row=1, max_row=len(all_rows)+1, max_col=7):
                    for cell in row:
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

            st.success("✅ Đã tạo xong bảng tổng hợp sạch 100%!")
            st.dataframe(df_final)
            st.download_button("📥 Tải file Excel Tổng Hợp", output.getvalue(), "Tong_Hop_Chi_Tiet.xlsx")
