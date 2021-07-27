from apply_password import *
from flask import *
from werkzeug.utils import secure_filename
import shutil
import pandas as pd
from table_extract import auto_table_extract
import webbrowser
from pathlib import Path
import os

cwd = Path.cwd()
ate = cwd.__str__()
flag = 0
counter = 0
count = 0
session = 0
export = 0
alpha = 0
beta = 0
upload_folder = ate + "\\upload"
excel_folder = ate + "\\excel"
csv_folder = ate + "\\csv"

shutil.rmtree(upload_folder)
os.makedirs(upload_folder)

if "output.xlsx" in os.listdir(excel_folder):
    shutil.rmtree(csv_folder)
    os.makedirs(csv_folder)
    UPLOAD_FOLDER = upload_folder
    app = Flask(__name__)
    app.secret_key = "secret key"
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
    app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
    ALLOWED_EXTENSIONS = set(['pdf'])

else:
    shutil.rmtree(csv_folder)
    os.makedirs(csv_folder)
    UPLOAD_FOLDER = upload_folder
    app = Flask(__name__)
    app.secret_key = "secret key"
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
    app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
    ALLOWED_EXTENSIONS = set(['pdf'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    flash("Upload a File!")
    return render_template('start.html')


@app.route('/uploads')
def upload_form():
    return render_template('Index.html')


@app.route('/how_to_use')
def how_to_use():
    return render_template("use.html")


@app.route('/uploads', methods=['POST'])
def upload_file():
    global a
    global flag
    global session
    global export
    global count
    flag = 0
    session = 1
    export = 0
    count = 0
    if request.method == 'POST':
        file = request.files['file']
        if file.filename == '':
            flash('Please Upload a File!')
            return redirect("/uploads")
        if allowed_file(file.filename):
            a = file.filename
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            flash('File Successfully Uploaded!')
            return redirect('/uploads')
        else:
            flash('ONLY PDF FILES ARE ALLOWED!')
            return redirect("/uploads")


@app.route('/pdf_view/')
def pdf_view():
    global session
    session = session
    if session == 1:
        import webbrowser as wb
        try:
            y = a.replace(" ", "_")
            filename = upload_folder + "\\" + y
            wb.open(filename)
            return redirect("/uploads")
        except:
            flash("Please Upload a File to view!")
            return redirect("/uploads")
    else:
        flash("New Session! Please Upload a File to view!")
        return redirect("/uploads")


@app.route('/pdf/<filename>')
def pdf_viewer(filename):
    arr = os.listdir(upload_folder)
    filename = arr[0]
    path = upload_folder + "//" + filename
    return send_file(path, attachment_filename=filename)


@app.route('/download/')
def download():
    global alpha
    global flag
    global counter
    global session
    global export
    global beta
    global count
    count = count
    beta = beta
    alpha = alpha
    export = export
    session = session
    flag1 = flag
    if session == 1 or count == 1:
        if flag1 == 1 or alpha == 1:
            session = 0
            flag = 0
            if export == 1:
                export = 0
            else:
                export = 1
            alpha = 0
            beta = 1
            try:
                y = a.replace(" ", "_")
                import re
                z = y.replace(".pdf", ".xlsx")
                filename = upload_folder + "\\" + z
                path = excel_folder + "\\output.xlsx"
                return send_file(path, mimetype='pdf', attachment_filename=z, as_attachment=True)
            except:
                flash("Please Extract the tables")
                return redirect("/uploads")
        else:
            flash("Please Upload a file or Extract the tables")
            return redirect("/uploads")
    else:
        flash("New Session! please upload a file")
        return redirect("/uploads")


@app.route('/download_with_password/')
def download_with_password():
    global flag
    global counter
    global session
    global export
    global alpha
    global beta
    global count
    beta1 = beta
    export == export
    session = session
    flag1 = flag
    count = count
    alpha = alpha
    if session == 1 or export == 1:
        if flag1 == 1 or beta1 == 1:
            path = excel_folder + "\\output.xlsx"
            set_password(path, processed_text)
            arr = os.listdir(excel_folder)
            filename = arr[0]
            path = excel_folder + "\\" + filename
            flag == 0
            counter = 1
            session = 0
            alpha = 1
            beta = 0
            if count == 1:
                count = 0
            else:
                count = 1
            return send_file(path, mimetype='xlsx', attachment_filename=path, as_attachment=True)
        else:
            flash("Please Upload a File!")
            return redirect("/uploads")

    else:
        flash("New Session! Please Upload a File!")
        return redirect("/uploads")


@app.route('/extract_table/')
def table_extraction():
    global flag
    global session
    session1 = session
    if session1 == 1:
        if len(os.listdir(upload_folder)) == 0:
            flash("Please Upload a File!")
            return redirect("/uploads")
        else:
            try:
                y = a.replace(" ", "_")
                auto_table_extract(upload_folder + "\\" + y)
                flash("Your Excel File is Ready!")
                flag = 1
                return redirect("/uploads")
            except:
                flash("Please Upload a File!")
                return redirect("/uploads")
    else:
        flash("New session! Please Upload a File!")
        return redirect("/uploads")


@app.route('/view/')
def data_frame():
    global flag
    global session
    session1 = session
    flag1 = flag
    if session1 == 1:
        if flag1 == 1:
            if len(os.listdir(csv_folder)) == 0:
                flash("Please Extract Tables!")
                return redirect("/uploads")
            else:
                arr = os.listdir(csv_folder)
                filename = arr[0]
                url = '/preview/' + filename
                return redirect(url)
        else:
            flash("Please Extract the tables!")
            return redirect("/uploads")
    else:
        flash("New Session! upload a new file!")
        return redirect("/uploads")


@app.route('/preview/<filename>')
def data_frame_show(filename):
    arr = os.listdir(csv_folder)
    filename = arr[0]
    path = csv_folder + "\\" + filename
    df = pd.read_csv(path, encoding='latin1')
    return df.to_html(header="true", table_id="table")


processed_text = "1234"


@app.route('/password_apply', methods=['GET', 'POST'])
def password_apply():
    if len(os.listdir(csv_folder)) == 0:
        flash("Please Extract Tables!")
        return redirect("/uploads")
    else:
        global processed_text
        if request.method == 'POST':
            text1 = request.form['CONFIRM_PASSWORD']
            processed_text = text1
        return render_template("password.html")


@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response


if __name__ == "__main__":
    FLASK_DEBUG = 1
    url = "http://127.0.0.1:5000/"
    webbrowser.open(url)
    app.run(threaded=True)
