import os
import comtypes.client
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
import uuid
import io

# --- 配置 ---
UPLOAD_FOLDER = 'uploads'
# 临时文件夹仍然需要，用于存放转换过程中生成的PDF
TEMP_FOLDER = 'temp_files'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['TEMP_FOLDER'] = TEMP_FOLDER
app.config['SECRET_KEY'] = 'supersecretkey'

# 确保上传和临时文件夹存在
for folder in [UPLOAD_FOLDER, TEMP_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_excel_to_pdf(excel_path, pdf_path):
    """将Excel文件转换为PDF，包含页面设置。"""
    excel = None
    workbook = None
    comtypes.CoInitialize()
    try:
        excel = comtypes.client.CreateObject("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(excel_path)

        for sheet in workbook.Sheets:
            try:
                page_setup = sheet.PageSetup
                page_setup.Orientation = 2
                page_setup.PaperSize = 9
                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1
                page_setup.FitToPagesTall = False
            except Exception:
                pass

        sheet_names = [sheet.Name for sheet in workbook.Sheets]
        workbook.Worksheets(sheet_names).Select()
        workbook.ActiveSheet.ExportAsFixedFormat(0, pdf_path)
        return True
    except Exception as e:
        print(f"转换时发生错误: {e}")
        return False
    finally:
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()
        comtypes.CoUninitialize()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('请求中没有文件部分')
        return redirect(url_for('index'))
    file = request.files['file']
    if file.filename == '':
        flash('没有选择文件')
        return redirect(url_for('index'))

    if file and allowed_file(file.filename):
        original_filename = secure_filename(file.filename)
        unique_id = uuid.uuid4().hex
        
        # 临时保存上传的Excel文件
        temp_excel_filename = f"{unique_id}_{original_filename}"
        temp_excel_path = os.path.join(app.config['UPLOAD_FOLDER'], temp_excel_filename)
        file.save(temp_excel_path)

        # 定义临时PDF输出路径
        pdf_filename_for_download = f"{os.path.splitext(original_filename)[0]}.pdf"
        temp_pdf_path = os.path.join(app.config['TEMP_FOLDER'], f"{unique_id}.pdf")

        # 执行转换
        success = convert_excel_to_pdf(os.path.abspath(temp_excel_path), os.path.abspath(temp_pdf_path))
        
        # 转换后立即删除上传的Excel文件
        os.remove(temp_excel_path)

        if not success:
            flash('文件转换失败，请检查服务器日志。')
            return redirect(url_for('index'))

        try:
            # 将生成的PDF读入内存缓冲区
            with open(temp_pdf_path, 'rb') as f:
                pdf_buffer = io.BytesIO(f.read())
            
            # 从服务器删除临时的PDF文件
            os.remove(temp_pdf_path)

            # 将缓冲区作为文件直接发送给客户端下载
            return send_file(
                pdf_buffer,
                as_attachment=True,
                download_name=pdf_filename_for_download,
                mimetype='application/pdf'
            )
        except FileNotFoundError:
            flash('读取转换后的文件时出错，请重试。')
            return redirect(url_for('index'))

    flash('文件类型不被允许')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True, port=8085)
