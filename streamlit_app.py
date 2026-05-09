import streamlit as st
import pdfplumber
import pandas as pd
import io

st.set_page_config(page_title="PDF to Excel Pro", layout="wide")
st.title("🚀 Tool PDF sang Excel (Chia đúng cột Table)")

uploaded_file = st.file_uploader("Tải file PDF báo cáo", type="pdf")

if uploaded_file:
    with st.spinner('Đang phân tích cấu trúc bảng...'):
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                for i, page in enumerate(pdf.pages):
                    # 1. Trích xuất bảng trước để lấy cấu trúc cột chuẩn
                    tables = page.extract_tables(table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_y_tolerance": 3,
                    })
                    
                    # 2. Trích xuất văn bản tự do
                    text = page.extract_text()
                    text_lines = text.split('\n') if text else []

                    # Tạo danh sách dữ liệu cuối cùng cho sheet này
                    final_page_data = []

                    # Đưa văn bản tự do vào (mỗi dòng là 1 list để nó nằm ở cột A)
                    for line in text_lines:
                        # Nếu dòng văn bản này trùng với nội dung trong bảng thì bỏ qua (tránh lặp)
                        final_page_data.append([line]) 

                    final_page_data.append(["--- BẢNG CHI TIẾT ---"])

                    # Đưa dữ liệu bảng vào (giữ nguyên danh sách các cột)
                    if tables:
                        for table in tables:
                            for row in table:
                                # Làm sạch dữ liệu từng ô
                                clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                                final_page_data.append(clean_row)

                    # 3. Chuyển vào DataFrame
                    # Cần đảm bảo DataFrame nhận diện đúng số cột lớn nhất
                    df_page = pd.DataFrame(final_page_data)
                    
                    # Lưu vào sheet
                    sheet_name = f"Trang {i+1}"
                    df_page.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        st.success(f"Đã xử lý {len(pdf.pages)} trang. Dữ liệu đã được chia cột!")

        st.download_button(
            label="📥 Tải file Excel (Đã chia cột)",
            data=output.getvalue(),
            file_name="Bao_cao_dung_dinh_dang.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
