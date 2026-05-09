import streamlit as st
import pdfplumber
import pandas as pd
import io

st.title("Tool Chuyển PDF Sang Excel - Chương Dương")

uploaded_file = st.file_uploader("Chọn file PDF của bạn", type="pdf")

if uploaded_file is not None:
    with pdfplumber.open(uploaded_file) as pdf:
        all_data = []
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table)
                all_data.append(df)
        
        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            st.write("Dữ liệu đã trích xuất:")
            st.dataframe(final_df)

            # Tạo nút tải về
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False)
            
            st.download_button(
                label="Tải file Excel về máy",
                data=output.getvalue(),
                file_name="ket_qua_chuyen_doi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
