import datetime
import openpyxl 
import subprocess
import os
from flask import Flask, flash, request, redirect, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]iasdfffsd/'

ALLOWED_EXTENSIONS = set(['xlsx','txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'csv',])



app.config['UPLOAD_FOLDER']='C:/Users/LENOVO/backend_file'#path where the uploaded file should be stored



@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        # submit an empty part without filename
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
    
        filename = secure_filename(file.filename)
            
        file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))

        path = u"C:\\Users\\LENOVO\\backend_file\\"+str(filename)
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj.active
        cell_obj = sheet_obj.cell(row = 2, column = 2)
        return cell_obj.value #return value in that cell in the excel sheet
    




if __name__ == '__main__':
      app.secret_key = 'super secret key'
      app.config['SESSION_TYPE'] = 'filesystem'
      app.debug = True
      app.run()