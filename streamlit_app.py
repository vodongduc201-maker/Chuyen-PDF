import streamlit as st
import pdfplumber
import pandas as pd
import io

# Cấu hình giao diện
st.set_page_config(page_title="Chương Dương PDF Automation", layout="wide")

st.title("🏭 Tool Chuyển Đổi Báo Cáo PDF sang Excel")
st.markdown("---")

# Hướng dẫn nhanh
st.sidebar.header("Hướng dẫn")
st.sidebar.info("""
1. Tải file PDF báo cáo lên.
2. Tool sẽ tự động nhận diện văn bản và bảng biểu.
3. Mỗi trang PDF sẽ được tách thành 1 Sheet Excel riêng.
""")

uploaded_file = st.file_uploader("Chọn file PDF (Nhiều trang)", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý dữ liệu... Vui lòng đợi trong giây lát.'):
        output = io.BytesIO()
        
        # Mở luồng ghi Excel chuyên sâu
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            with pdfplumber.open(uploaded_file) as pdf:
                total_pages = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    # 1. Trích xuất bảng với thuật toán nhận diện đường kẻ
                    # Điều này giúp chia cột A, B, C... chính xác như Table Excel
                    tables = page.extract_tables(table_settings={
                        "vertical_strategy": "lines",
                        "horizontal_strategy": "lines",
                        "snap_y_tolerance": 3,
                    })
                    
                    # 2. Trích xuất văn bản tự do (thông tin Header/Footer)
                    text = page.extract_text()
                    lines = text.split('\n') if text else []

                    # Tạo danh sách chứa dữ liệu của trang hiện tại
                    current_page_rows = []

                    # Đưa văn bản tự do vào (Cố định ở cột đầu tiên)
                    for line in lines:
                        current_page_rows.append([line]) 

                    current_page_rows.append([""]) # Dòng trống ngăn cách
                    current_page_rows.append(["--- BẢNG CHI TIẾT ---"])

                    # 3. Đưa bảng vào (Tách cột)
                    if tables:
                        for table in tables:
                            for row in table:
                                # Làm sạch dữ liệu: bỏ ký tự xuống dòng trong ô để không bị lỗi format
                                clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                                current_page_rows.append(clean_row)

                    # 4. Tạo DataFrame và lưu vào Sheet tương ứng
                    df = pd.DataFrame(current_page_rows)
                    sheet_name = f"Trang {i+1}"
                    df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        
        st.success(f"✅ Hoàn thành! Đã xử lý tổng cộng {total_pages} trang.")

        # Nút tải xuống file kết quả
        st.download_button(
            label="📥 Tải file Excel (.xlsx)",
            data=output.getvalue(),
            file_name="Bao_cao_tong_hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
