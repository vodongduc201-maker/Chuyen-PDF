import streamlit as st
import pdfplumber
import pandas as pd
import io
import re
from openpyxl.styles import Border, Side

st.set_page_config(page_title="PDF Master Summary", layout="wide")
st.title("📊 Master Tool: Lọc sạch tên Store (Cột D)")

uploaded_file = st.file_uploader("Tải file PDF của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý loại bỏ thông tin dư thừa...'):
        all_rows = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            # Duyệt từ trang 2 đến hết
            for i in range(1, len(pdf.pages)):
                page = pdf.pages[i]
                
                # 1. Lấy thông tin Header bằng văn bản
                text = page.extract_text()
                so_match = re.search(r"SO number:\s*(\d+)", text)
                order_date = re.search(r"Order date:\s*([\d/ :]+)", text)
                delivery_date = re.search(r"Delivery date:\s*([\d/]+)", text)
                
                so = so_match.group(1) if so_match else ""
                o_date = order_date.group(1).split()[0] if order_date else ""
                d_date = delivery_date.group(1) if delivery_date else ""

                # 2. LẤY TÊN STORE (CỘT D) CHÍNH XÁC THEO TỌA ĐỘ
                # Chúng ta sẽ tìm chữ "For Store" và lấy nội dung ngay bên dưới nó
                store_name = ""
                words = page.extract_words()
                for idx, word in enumerate(words):
                    if "For" in word['text'] and idx + 1 < len(words) and "Store" in words[idx+1]['text']:
                        # Tìm các từ nằm ở dòng ngay dưới cụm "For Store" trong cùng khu vực dọc (x)
                        target_y = word['bottom']
                        target_x0 = word['x0'] - 10
                        target_x1 = word['x1'] + 50
                        
                        store_words = [w['text'] for w in words if 
                                       target_y < w['top'] < target_y + 25 and 
                                       target_x0 < w['x0'] < target_x1 + 100]
                        
                        candidate = " ".join(store_words).strip()
                        # Loại bỏ các từ khóa không mong muốn
                        if "CTY TNHH" not in candidate and "LOGISTIC" not in candidate:
                            store_name = candidate
                        break

                # 3. Trích xuất bảng sản phẩm
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                        article_id = clean_row[0].strip()
                        
                        if article_id.isdigit() and len(article_id) >= 10:
                            all_rows.append({
                                "SO Number": so,
                                "Order Date": o_date,
                                "Delivery Date": d_date,
                                "Store": store_name, # Cột D đã được lọc sạch
                                "Article": article_id,
                                "Description": clean_row[1],
                                "OU Qty": clean_row[5] if len(clean_row) > 5 else ""
                            })

        if all_rows:
            df_final = pd.DataFrame(all_rows)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name="Master Summary")
                ws = writer.sheets["Master Summary"]
                thin = Side(style='thin')
                for row in ws.iter_rows(min_row=1, max_row=len(all_rows)+1, max_col=7):
                    for cell in row:
                        cell.border = Border(top=thin, left=thin, right=thin, bottom=thin)

            st.success("✅ Đã xử lý xong! Cột D đã loại bỏ thông tin CTY TNHH DV EB và Kho vận.")
            st.dataframe(df_final)
            st.download_button("📥 Tải file Excel Tổng Hợp", output.getvalue(), "Ket_Qua_Tong_Hop.xlsx")
