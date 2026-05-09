import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Multi-Sheet", layout="wide")

st.title("🚀 Tool PDF sang Excel (Mỗi trang 1 Sheet)")
st.info("Dữ liệu bao gồm văn bản tự do và bảng biểu")

uploaded_file = st.file_uploader("Tải file PDF của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý từng trang...'):
        output = io.BytesIO()
        
        # Mở công cụ ghi Excel
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_data = []
                    
                    # 1. Lấy văn bản tự do (ngoài bảng)
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            page_data.append([line]) # Mỗi dòng văn bản là 1 hàng
                    
                    page_data.append(["--- BẢNG DỮ LIỆU ---"])
                    
                    # 2. Lấy các bảng (giữ cấu trúc cột)
                    tables = page.extract_tables(table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_y_tolerance": 3,
                    })
                    
                    for table in tables:
                        for row in table:
                            clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                            page_data.append(clean_row)
                    
                    # 3. Chuyển dữ liệu trang hiện tại vào 1 Sheet riêng
                    if page_data:
                        df_page = pd.DataFrame(page_data)
                        # Đặt tên Sheet theo trang (Trang 1, Trang 2...)
                        sheet_name = f"Trang {i+1}"
                        df_page.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        st.success(f"Đã xử lý xong {len(pdf.pages)} trang!")

        # Nút tải file
        st.download_button(
            label="📥 Tải file Excel (Nhiều Sheet)",
            data=output.getvalue(),
            file_name="Bao_cao_theo_trang.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
