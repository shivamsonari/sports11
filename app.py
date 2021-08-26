from flask import Flask, render_template, redirect
import requests
import json
import os
import pandas as pd
import xlrd
from flask import request
import zipfile
import shutil
from flask import send_file


app = Flask(__name__)
cache= {}


@app.route("/")
@app.route("/home")
def home():
    return render_template("home.html")

@app.route("/generate", methods = ["GET", "POST"])
def generate():
    if request.method == "POST":
        sheet = request.form.getlist('sheet')
        
        
        inputFile = cache['file']

#getting sheet names
        xls = xlrd.open_workbook(inputFile, on_demand=True)
        sheet_names = xls.sheet_names()

        path = "static/"
        zipf = zipfile.ZipFile('Name.zip','w', zipfile.ZIP_DEFLATED)

        #create a new excel file for every sheet
        for name in sheet_names:
            
            if name in sheet :
                
    
                parsing = pd.ExcelFile(inputFile).parse(sheetname = name)

                #writing data to the new excel file
                parsing.to_excel(path+str(name)+".xlsx", index=False)
                zipf.write(path+str(name)+".xlsx")
                
                print(request)
    zipf.close()
    return send_file('Name.zip',
            mimetype = 'zip',
            attachment_filename= 'Name.zip',
            as_attachment = True)


    

@app.route('/uploader', methods = ['GET', 'POST'])

def upload_file():
   if request.method == 'POST':
      f = request.files['file']
      if 'file' in cache.keys():
          
          os.remove(cache['file'])
          shutil.rmtree("static")
          os.mkdir("static")
      cache['file']=f.filename
      f.save(f.filename)
      return redirect("/")

     



if __name__ == '__main__':
    os.environ['FLASK_ENV'] = 'development'
    app.run(debug=True)
    