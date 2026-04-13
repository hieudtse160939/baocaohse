import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import io
import zipfile
import os
import subprocess
import tempfile

st.set_page_config(page_title="Hệ thống Phiếu Điểm - Hoa Sen School", layout="wide")

# Giao diện chính
st.title("🎓 Công cụ Quản lý Phiếu Điểm")
st.info("Hệ thống dành riêng cho cán bộ giáo viên Trường TH, THCS và THPT Hoa Sen.")

# --- PHẦN 1: TẢI FILE MẪU ---
with st.sidebar:
    st.header("Hướng dẫn & Tài liệu")
    # Giả sử bạn đã up file 'template_excel.xlsx' lên GitHub cùng thư mục
    try:
        with open("template_excel.xlsx", "rb") as f:
            st.download_button(
                label="📥 Tải File Excel Mẫu",
                data=f,
                file_name="Mau_Danh_Sach_Diem_HoaSen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except FileNotFoundError:
        st.warning("Vui lòng upload file 'template_excel.xlsx' lên GitHub.")

# --- PHẦN 2: UPLOAD VÀ XỬ LÝ ---
uploaded_file = st.file_uploader("Bước 1: Tải lên file Excel điểm", type=["xlsx"])
export_format = st.radio("Bước 2: Chọn định dạng tải về", ["PDF (Khuyên dùng)", "Word (DOCX)"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    df.columns = [col.replace(' ', '_').replace('\n', '_') for col in df.columns]
    st.write("Xem trước dữ liệu:", df.head(5))
    
    if st.button("🚀 Bắt đầu tạo phiếu điểm"):
        zip_buffer = io.BytesIO()
        template_path = "template.docx" # File Word mẫu của bạn
        
        if not os.path.exists(template_path):
            st.error("Không tìm thấy file 'template.docx'. Hãy kiểm tra lại trên GitHub.")
        else:
            with st.spinner("Đang xử lý... Vui lòng đợi trong giây lát."):
                with zipfile.ZipFile(zip_buffer, "w") as zf:
                    for index, row in df.iterrows():
                        if pd.isna(row.get('Họ_và_Tên')): continue
                        
                        # Tạo file Word trong thư mục tạm
                        with tempfile.TemporaryDirectory() as tmp_dir:
                            doc = DocxTemplate(template_path)
                            
                            # Định dạng số trước khi render
                            context = row.to_dict()
                            # (Tùy chỉnh: thêm logic định dạng 2 chữ số thập phân tại đây)
                            
                            doc.render(context)
                            
                            file_name_base = f"{row['Họ_và_Tên']} - {row.get('Lớp', 'Lớp')}"
                            docx_path = os.path.join(tmp_dir, f"{file_name_base}.docx")
                            doc.save(docx_path)
                            
                            if export_format == "PDF (Khuyên dùng)":
                                # Sử dụng LibreOffice để chuyển Word sang PDF
                                try:
                                    subprocess.run([
                                        "libreoffice", "--headless", "--convert-to", "pdf",
                                        "--outdir", tmp_dir, docx_path
                                    ], check=True)
                                    
                                    pdf_path = os.path.join(tmp_dir, f"{file_name_base}.pdf")
                                    with open(pdf_path, "rb") as f:
                                        zf.writestr(f"{file_name_base}.pdf", f.read())
                                except Exception as e:
                                    st.error(f"Lỗi khi tạo PDF cho {file_name_base}: {e}")
                            else:
                                # Chỉ lưu file Word
                                with open(docx_path, "rb") as f:
                                    zf.writestr(f"{file_name_base}.docx", f.read())
                
                st.success("Đã hoàn thành tất cả phiếu điểm!")
                st.download_button(
                    label="📥 Tải về file ZIP kết quả",
                    data=zip_buffer.getvalue(),
                    file_name=f"Ket_Qua_Phieu_Diem_{export_format}.zip",
                    mime="application/zip"
                )
