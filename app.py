from flask import Flask, render_template, request, send_file
import os
import getting_data, converter, shorten

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 1. Nhận file PDF từ người dùng
        uploaded_pdf = request.files['file_tkb']
        if uploaded_pdf.filename != '':
            pdf_path = os.path.join(UPLOAD_FOLDER, "All_TKB.pdf")
            uploaded_pdf.save(pdf_path)
            
            # Đường dẫn các file đầu ra
            doc_out = os.path.join(OUTPUT_FOLDER, "TKB.docx")
            pdf_out = os.path.join(OUTPUT_FOLDER, "TKB.pdf")
            sample_doc = "Sample_TKB.docx" # File mẫu để sẵn ở thư mục gốc
            
            try:
                # 2. Xử lý dữ liệu
                getting_data.start(pdf_path, sample_doc, doc_out)
                converter.start(OUTPUT_FOLDER, doc_out, pdf_out)
                tkb_rut_gon = shorten.start(doc_out) # Nhớ sửa shorten.py trả về text
                
                # Trả về trang kết quả kèm văn bản TKB rút gọn
                return render_template('index.html', result_text=tkb_rut_gon, success=True)
            except Exception as e:
                return render_template('index.html', error=str(e))
                
    return render_template('index.html')

# API để tải file về
@app.route('/download/<file_type>')
def download(file_type):
    if file_type == 'docx':
        return send_file(os.path.join(OUTPUT_FOLDER, "TKB.docx"), as_attachment=True)
    elif file_type == 'pdf':
        return send_file(os.path.join(OUTPUT_FOLDER, "TKB.pdf"), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)