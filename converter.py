import subprocess
import fitz
import os

def pdf_to_png(location, sel_file):
    pdf_document = fitz.open(sel_file)
    pix = pdf_document[0].get_pixmap(dpi=600)
    pix.save(os.path.join(location, "TKB.png"))
    pdf_document.close()

def start(location, sel_docx_file, sel_pdf_file):
    # Lệnh gọi LibreOffice chạy ngầm, không mở giao diện (--headless)
    # --outdir chỉ định thư mục đích để lưu file PDF
    command = [
        'libreoffice', '--headless', '--convert-to', 'pdf',
        '--outdir', location, sel_docx_file
    ]
    
    try:
        # Thực thi lệnh
        subprocess.run(command, check=True)
        
        # Gọi tiếp hàm chuyển từ PDF sang PNG (nếu bạn cần dùng tới TKB.png)
        pdf_to_png(location, sel_pdf_file)
        
    except subprocess.CalledProcessError as e:
        raise Exception(f"Lỗi khi chuyển đổi file bằng LibreOffice: {e}")