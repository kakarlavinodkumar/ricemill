from flask import Flask, render_template, request,url_for,redirect,send_from_directory
from werkzeug import secure_filename
from  excel import processExcel
import os
from io import BytesIO
import zipfile
import time

app = Flask(__name__)

app = Flask(__name__)


@app.route('/upload')
def upload_file_form():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
   if request.method == 'POST':
        #f = request.files['file']
        print request.files
        uploaded_files = request.files.getlist("file")
        print uploaded_files
        uploads =os.path.join(os.path.dirname(os.path.realpath(__file__)))   
        zipf = zipfile.ZipFile('Python.zip', 'w', zipfile.ZIP_DEFLATED)

        for f in uploaded_files :
            f.save(secure_filename(f.filename))
            print "File Name  "+secure_filename(f.filename)
            processExcel(secure_filename(f.filename))
            splited = secure_filename(f.filename).split(".")
            dest_filename = splited[0]+"output"+".xlsx"
            zipf.write(dest_filename)

        zipf.close()
        return render_template("uploader.html")

@app.route('/download', methods=['GET', 'POST'])
def download():
    uploads =os.path.join(os.path.dirname(os.path.realpath(__file__)))   
    #render_template("upload.html")
    return send_from_directory(directory=uploads, filename="Python.zip")
    
@app.route('/')
def home_page():
    return render_template('upload.html')

@app.route("/create_zip")
def createZip() :
    
    uploads =os.path.join(os.path.dirname(os.path.realpath(__file__)))   
    files = ["output.xlsx","styles2.py"]
    zipf = zipfile.ZipFile('Python.zip', 'w', zipfile.ZIP_DEFLATED)
    
    for file in files:
        zipf.write(file)
    zipf.close()

    return send_from_directory(directory=uploads, filename="Python.zip")


    
if __name__ == '__main__':
   app.run()