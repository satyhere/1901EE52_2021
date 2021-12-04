from flask import Flask, render_template, request, redirect, url_for
import os
from os.path import join, dirname, realpath

app = Flask(__name__)


UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] =  UPLOAD_FOLDER



@app.route('/')
def index():
    
    return render_template('main.html')
@app.route("/", methods=['POST'])
def uploadFiles():
      
      uploaded_file = request.files['file']
      if uploaded_file.filename != '':
           file_path = os.path.join(app.config['UPLOAD_FOLDER'], uploaded_file.filename)
          
           uploaded_file.save(file_path)
          
      return redirect(url_for('main.html'))

if (__name__ == "__main__"):
    app.run(port = 5000)

