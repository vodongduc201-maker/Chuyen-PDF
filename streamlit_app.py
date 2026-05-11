import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl.styles import Border, Side

st.set_page_config(page_title="PDF Master Summary", layout="wide")
st.title("📊 Master Tool: Fix lỗi Cột D (Store Name)")

uploaded_file = st.file_uploader("Tải file PDF của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang nhặt dữ liệu...'):
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            # Duyệt từ trang 2 đến hết
            for i in range(1, len(pdf.pages)):
                page = pdf.pages[i]
                text = page.extract_text()
                
                # 1. Lấy SO, Order Date, Delivery Date
                so_match = re.search(r"SO number:\s*(\d+)", text)
                order_date = re.search(r"Order date:\s*([\d/ :]+)", text)
                delivery_date = re.search(r"Delivery date:\s*([\d/]+)", text)
                
                so = so_match.group(1) if so_match else ""
                o_date = order_date.group(1).split()[0] if order_date else ""
                d_date = delivery_date.group(1) if delivery_date else ""

                # 2. LẤY TÊN STORE (Cột D) - Cách mới ổn định hơn
                store = ""
                lines = text.split('\n')
                for idx, line in enumerate(lines):
                    if "For Store" in line:
                        # Thường tên Store nằm ở 1 hoặc 2 dòng ngay sau chữ "For Store"
                        if idx + 1 < len(lines):
                            store = lines[idx + 1].strip()
                        break
                
                # 3. Trích xuất bảng sản phẩm
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                        
                        # Điều kiện lấy dòng sản phẩm: Cột 1 là số (Article)
                        article_id = clean_row[0].strip()
                        if article_id.isdigit() and len(article_id) >= 10:
                            all_rows.append({
                                "SO Number": so,
                                "Order Date": o_date,
                                "Delivery Date": d_date,
                                "Store": store, # Thông tin cột D
                                "Article": article_id,
                                "Description": clean_row[1],
                                "OU Qty": clean_row[5] if len(clean_row) > 5 else ""
                            })

        if all_rows:
            df_final = pd.DataFrame(all_rows)
            
            # Xuất file Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Master Summary")
                
                # Đóng khung
                ws = writer.sheets["Master Summary"]
                thin = Side(style='thin')
                for row in ws.iter_rows(min_row=1, max_row=len(all_rows)+1, max_col=7):
                    for cell in row:
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

            st.success("✅ Đã xử lý xong! Cột D đã có đầy đủ tên Store.")
            st.dataframe(df_final) # Hiển thị xem trước để bạn kiểm tra cột D
            st.download_button("📥 Tải file Excel Tổng Hợp", output.getvalue(), "Tong_Hop_Master.xlsx")
