import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF Automation Tool", layout="wide")

st.title("🚀 Tool Chuyển Đổi Đặc Thù (Trang 1 + Excel Tabs)")
st.info("Trang 1 sẽ giữ cấu trúc văn bản, các trang sau sẽ tách bảng vào từng Sheet.")

uploaded_file = st.file_uploader("Tải file PDF 29 trang của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý dữ liệu...'):
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                for i, page in enumerate(pdf.pages):
                    page_data = []
                    
                    # XỬ LÝ TRANG 1: Giữ nguyên cấu trúc văn bản
                    if i == 0:
                        text = page.extract_text()
                        if text:
                            for line in text.split('\n'):
                                page_data.append([line])
                    
                    # XỬ LÝ CÁC TRANG CÒN LẠI: Tách bảng chuyên sâu
                    else:
                        # Sử dụng thuật toán nhận diện đường kẻ để chia cột chuẩn
                        tables = page.extract_tables(table_settings={
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "snap_y_tolerance": 3,
                            "intersection_x_tolerance": 3,
                        })
                        
                        if tables:
                            for table in tables:
                                for row in table:
                                    # Làm sạch ô dữ liệu
                                    clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                                    page_data.append(clean_row)
                        else:
                            # Nếu trang đó không có đường kẻ rõ ràng, dùng text thô
                            text = page.extract_text()
                            if text:
                                for line in text.split('\n'):
                                    page_data.append([line])

                    # Ghi vào Sheet tương ứng
                    if page_data:
                        df = pd.DataFrame(page_data)
                        sheet_name = f"Trang {i+1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

        st.success(f"✅ Đã xử lý xong {len(pdf.pages)} trang!")

        st.download_button(
            label="📥 Tải file Excel kết quả",
            data=output.getvalue(),
            file_name="Bao_cao_tach_trang_3.4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
