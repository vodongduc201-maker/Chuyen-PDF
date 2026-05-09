import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Custom", layout="wide")
st.title("🚀 Tool PDF - Tự động xóa ô A2, A3, B2, B3, D2, D3")

uploaded_file = st.file_uploader("Tải file PDF của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý và làm sạch dữ liệu...'):
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_data = []
                    
                    # 1. Trích xuất văn bản và bảng
                    text = page.extract_text()
                    if text:
                        for line in text.split('\n'):
                            page_data.append([line])
                    
                    tables = page.extract_tables(table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_y_tolerance": 3,
                    })
                    
                    if tables:
                        for table in tables:
                            for row in table:
                                clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                                page_data.append(clean_row)

                    if page_data:
                        # Chuyển thành DataFrame
                        df = pd.DataFrame(page_data)

                        # 2. XỬ LÝ BỎ CÁC Ô THEO YÊU CẦU (Chỉ áp dụng từ Trang 2 trở đi hoặc tất cả)
                        # Lưu ý: Index trong lập trình bắt đầu từ 0 (A=0, B=1, C=2, D=3...)
                        # Hàng 2 trong Excel là index 1, Hàng 3 là index 2
                        
                        locations_to_clear = [
                            (1, 0), (2, 0), # A2, A3
                            (1, 1), (2, 1), # B2, B3
                            (1, 3), (2, 3)  # D2, D3
                        ]

                        for row_idx, col_idx in locations_to_clear:
                            try:
                                # Kiểm tra nếu DataFrame có đủ hàng và cột đó thì mới xóa
                                if row_idx < len(df) and col_idx < len(df.columns):
                                    df.iat[row_idx, col_idx] = ""
                            except:
                                pass

                        # 3. Ghi vào Sheet
                        sheet_name = f"Trang {i+1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        st.success("✅ Đã xử lý xong! Các ô A2, A3, B2, B3, D2, D3 đã được để trống.")

        st.download_button(
            label="📥 Tải file Excel đã làm sạch",
            data=output.getvalue(),
            file_name="Bao_cao_clean.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
