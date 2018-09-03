from flask import Flask, render_template, request,url_for,redirect,send_from_directory
from werkzeug import secure_filename
from  excel import processExcel
import os

app = Flask(__name__)

app = Flask(__name__)


@app.route('/upload')
def upload_file_form():
   return render_template('upload.html')
	
@app.route('/uploader', methods = ['GET', 'POST'])
def upload_file():
   if request.method == 'POST':
      f = request.files['file']
      f.save(secure_filename(f.filename))
      print "File Name  "+secure_filename(f.filename)
      processExcel(secure_filename(f.filename))
      return render_template("uploader.html")

@app.route('/download', methods=['GET', 'POST'])
def download():
    uploads =os.path.join(os.path.dirname(os.path.realpath(__file__)))   
    render_template("upload.html")
    return send_from_directory(directory=uploads, filename="output.xlsx")
    
@app.route('/')
def home_page():
    return render_template('upload.html')

if __name__ == '__main__':
   app.run()