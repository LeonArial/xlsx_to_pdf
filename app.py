import os
import comtypes.client
from flask import Flask, request, render_template, send_from_directory, flash, redirect, url_for
from werkzeug.utils import secure_filename
import uuid

# --- 配置 ---
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER
app.config['SECRET_KEY'] = 'supersecretkey'

# 确保上传和下载文件夹存在
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
if not os.path.exists(DOWNLOAD_FOLDER):
    os.makedirs(DOWNLOAD_FOLDER)


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert_excel_to_pdf(excel_path, pdf_path):
    """将Excel文件转换为PDF，包含页面设置。"""
    excel = None
    workbook = None
    # 在Web环境中，为每个线程初始化COM库
    comtypes.CoInitialize()
    try:
        excel = comtypes.client.CreateObject("Excel.Application")
        excel.Visible = False
        
        workbook = excel.Workbooks.Open(excel_path)

        # 设置页面布局
        for sheet in workbook.Sheets:
            try:
                page_setup = sheet.PageSetup
                page_setup.Orientation = 2  # 横向
                page_setup.PaperSize = 9    # A4
                page_setup.Zoom = False
                page_setup.FitToPagesWide = 1
                page_setup.FitToPagesTall = False
            except Exception:
                # 忽略无法设置页面布局的工作表
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
        # 释放COM库
        comtypes.CoUninitialize()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('请求中没有文件部分')
        return redirect(request.url)
    file = request.files['file']
    if file.filename == '':
        flash('没有选择文件')
        return redirect(request.url)
    if file and allowed_file(file.filename):
        # 使用安全的文件名，并添加唯一前缀以避免冲突
        original_filename = secure_filename(file.filename)
        unique_id = uuid.uuid4().hex
        input_filename = f"{unique_id}_{original_filename}"
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(excel_path)

        # 定义输出的PDF路径
        output_filename = f"{os.path.splitext(original_filename)[0]}.pdf"
        pdf_path = os.path.join(app.config['DOWNLOAD_FOLDER'], output_filename)

        # 执行转换
        success = convert_excel_to_pdf(os.path.abspath(excel_path), os.path.abspath(pdf_path))

        # 清理上传的Excel文件
        os.remove(excel_path)

        if success:
            # 提供下载链接
            return redirect(url_for('download_file', filename=output_filename))
        else:
            flash('文件转换失败，请检查服务器日志。')
            return redirect(url_for('index'))

    flash('文件类型不被允许')
    return redirect(url_for('index'))

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename, as_attachment=True)

if __name__ == '__main__':
    # 使用 waitress 或其他生产级WSGI服务器部署，而不是Flask自带的开发服务器
    # from waitress import serve
    # serve(app, host="0.0.0.0", port=5000)
    app.run(debug=True) # 仅用于本地开发测试
