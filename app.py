import streamlit as st
import pandas as pd
import openpyxl
import io
import re
import os
import zipfile
from docxtpl import DocxTemplate

# =====================================================================
# CONFIG & STYLE
# =====================================================================
st.set_page_config(page_title="HSE Admin Tool - Hoa Sen School", page_icon="🏫", layout="wide")

# Hàm quay lại Menu chính
def back_to_menu():
    st.session_state.page = "Menu"

# Khởi tạo trạng thái trang nếu chưa có
if 'page' not in st.session_state:
    st.session_state.page = "Menu"

# =====================================================================
# LOGIC APP 1: THỐNG KÊ TKB (Dựa trên code bạn cung cấp)
# =====================================================================
# [Copy các hàm TU_DIEN, get_standard_name, phan_loai_khoi, process_tkb_data của bạn vào đây]
# (Để code gọn, mình sẽ giả định các hàm này đã được định nghĩa phía trên)

def app_thong_ke_tkb():
    st.header("📊 Thống kê Thời khóa biểu")
    if st.button("⬅️ Quay lại Menu"): back_to_menu()
    
    uploaded_file = st.file_uploader("Tải lên file TKB (.xlsx)", type=["xlsx"], key="tkb_up")
    if uploaded_file:
        # Sử dụng hàm process_tkb_data bạn đã viết
        # df_ket_qua = process_tkb_data(uploaded_file)
        # ... (Hiển thị dataframe và bộ lọc như code cũ của bạn)
        st.info("Tính năng Thống kê TKB đang sẵn sàng...")

# =====================================================================
# LOGIC APP 2: XUẤT PHIẾU ĐIỂM (MAIL MERGE)
# =====================================================================
def app_xuat_phieu_diem():
    st.header("🎓 Xuất Phiếu Điểm Tự Động")
    if st.button("⬅️ Quay lại Menu"): back_to_menu()
    
    col1, col2 = st.columns(2)
    with col1:
        template_word = st.file_uploader("1. Tải lên mẫu Word (template.docx)", type=["docx"])
    with col2:
        data_excel = st.file_uploader("2. Tải lên file điểm (Excel)", type=["xlsx"])
    
    if template_word and data_excel:
        df = pd.read_excel(data_excel)
        df.columns = [col.replace(' ', '_').replace('\n', '_') for col in df.columns]
        
        if st.button("🚀 Bắt đầu trộn dữ liệu"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for _, row in df.iterrows():
                    if pd.isna(row.get('Họ_và_Tên')): continue
                    doc = DocxTemplate(template_word)
                    context = row.to_dict()
                    # Logic format số 2 chữ số thập phân
                    for key, val in context.items():
                        if isinstance(val, (int, float)) and not pd.isna(val):
                            context[key] = f"{val:.2f}"
                    
                    doc.render(context)
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    zf.writestr(f"{row['Họ_và_Tên']} - {row.get('Lớp', 'Lớp')}.docx", doc_io.getvalue())
            
            st.success("Đã tạo xong!")
            st.download_button("📥 Tải về file ZIP (.docx)", data=zip_buffer.getvalue(), file_name="Phieu_Diem.zip")

# =====================================================================
# LOGIC APP 3: HƯỚNG DẪN & FILE MẪU
# =====================================================================
def app_huong_dan():
    st.header("📂 Kho File Mẫu & Hướng Dẫn")
    if st.button("⬅️ Quay lại Menu"): back_to_menu()
    
    st.write("Tải các file mẫu dưới đây để đảm bảo hệ thống chạy chính xác:")
    # Giả sử các file này nằm cùng thư mục trên GitHub
    # st.download_button("Tải File Excel Mẫu Điểm", data=..., file_name="Mau_Diem.xlsx")
    st.markdown("""
    1. **Mẫu TKB:** File xuất từ phần mềm xếp lịch của trường.
    2. **Mẫu Phiếu điểm:** Các trường dữ liệu phải nằm trong `{{ }}`.
    3. **Lưu ý:** Tên cột trong Excel phải khớp với tên trong file Word.
    """)

# =====================================================================
# MAIN ROUTING (ĐIỀU HƯỚNG)
# =====================================================================
if st.session_state.page == "Menu":
    st.title("🏫 HSE Digital Hub - Hoa Sen School")
    st.subheader("Chào Hiếu, vui lòng chọn công cụ cần làm việc:")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("### 📊 Thống kê TKB")
        st.write("Phân tích số tiết, lọc theo giáo viên/khối.")
        if st.button("Truy cập TKB"):
            st.session_state.page = "TKB"
            st.rerun()

    with col2:
        st.success("### 🎓 Xuất Phiếu Điểm")
        st.write("Trộn dữ liệu Excel vào mẫu Word hàng loạt.")
        if st.button("Truy cập Phiếu Điểm"):
            st.session_state.page = "PhieuDiem"
            st.rerun()

    with col3:
        st.warning("### 📂 File Mẫu")
        st.write("Tải template Excel và Word chuẩn.")
        if st.button("Xem hướng dẫn"):
            st.session_state.page = "HuongDan"
            st.rerun()

elif st.session_state.page == "TKB":
    app_thong_ke_tkb()
elif st.session_state.page == "PhieuDiem":
    app_xuat_phieu_diem()
elif st.session_state.page == "HuongDan":
    app_huong_dan()
