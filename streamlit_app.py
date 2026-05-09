import streamlit as st
import pdfplumber
import pandas as pd
import io
from openpyxl.styles import Border, Side

# Cấu hình giao diện
st.set_page_config(page_title="PDF Automation Tool", layout="wide")

st.title("🚀 Tool Chuyển Đổi Đặc Thù (Trang 1 + Excel Tabs)")
st.info("Trang 1 sẽ giữ cấu trúc văn bản, các trang sau sẽ tách bảng và ĐÓNG KHUNG ô.")

uploaded_file = st.file_uploader("Tải file PDF của bạn lên", type="pdf")

if uploaded_file:
    with st.spinner('Đang xử lý dữ liệu và đóng khung...'):
        output = io.BytesIO()
        
        # Định nghĩa kiểu viền (Border) - Nét mảnh, màu đen
        thin_border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
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
                        tables = page.extract_tables(table_settings={
                            "vertical_strategy": "lines",
                            "horizontal_strategy": "lines",
                            "snap_y_tolerance": 3,
                            "intersection_x_tolerance": 3,
                        })
                        
                        if tables:
                            for table in tables:
                                for row in table:
                                    clean_row = [str(cell).replace('\n', ' ') if cell else "" for cell in row]
                                    page_data.append(clean_row)
                        else:
                            text = page.extract_text()
                            if text:
                                for line in text.split('\n'):
                                    page_data.append([line])

                    # GHI VÀO EXCEL VÀ ĐÓNG KHUNG
                    if page_data:
                        df = pd.DataFrame(page_data)
                        sheet_name = f"Trang {i+1}"
                        df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                        
                        # Truy cập vào sheet vừa tạo để kẻ viền
                        worksheet = writer.sheets[sheet_name]
                        for row in worksheet.iter_rows(min_row=1, max_row=len(page_data), 
                                                       min_col=1, max_col=df.shape[1]):
                            for cell in row:
                                cell.border = thin_border

        st.success(f"✅ Đã xử lý xong {len(pdf.pages)} trang và đóng khung thành công!")

        st.download_button(
            label="📥 Tải file Excel có đóng khung",
            data=output.getvalue(),
            file_name="Bao_cao_co_khung.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
