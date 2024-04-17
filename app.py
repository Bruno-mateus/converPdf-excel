import os
from flask import Flask, flash, request, redirect,url_for
from werkzeug.utils import secure_filename
from flask import send_from_directory
from spire.pdf.common import *
from spire.pdf import *
import fitz
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.add_url_rule(
    "/uploads/<name>", endpoint="download_file", build_only=True
)
def convert_file(filename,pagesNum):
   
    pages=pagesNum

    
    pagesArr = pages.replace(','," ").split()
    pagesArrNum = []
    for page in pagesArr:
        pagesArrNum.append(int(page)-1)

    

    # Open the PDF file
    doc=fitz.open(f"uploads/{filename}")
    
    # # Say, you like to save the first 6 pages, first page is 0
    doc.select(pagesArrNum)
   
    # doc.pages
    # # Save the selected pages to a new PDF
    doc.save(f"uploads/temp/{filename}")

    # Create an object of the PdfDocument class
    pdf = PdfDocument()
     # Load a PDF file
    pdf.LoadFromFile(f"uploads/temp/{filename}")

    convertOptions = XlsxLineLayoutOptions(False, True, False, True, False)
    pdf.ConvertOptions.SetPdfToXlsxOptions(convertOptions)
    file = filename.replace(".pdf","")
    # # Save the PDF file to Excel XLSX format
    pdf.SaveToFile(f"uploads/{file}.xlsx", FileFormat.XLSX)
    # Close the PdfDocument object
    pdf.Close()



@app.route('/uploads/<name>')
def download_file(name):
    return send_from_directory(app.config["UPLOAD_FOLDER"], name)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        
        
        # check if the post request has the file part
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['file']
        pages = request.form['pages']
    
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            #flash('No selected file')
            return '''
    <!doctype html>
        <title>Upload new File</title>
        <strong style="color:red;">Selecione um arquivo</strong>
    <form method=post enctype=multipart/form-data style="margin-top:1rem">
      <input type=file name=file>
      <input type=text name=pages>
      <input type=submit value=Upload>
    </form>'''
        if pages == '':
            print(pages)
            #flash('No selected file')
            return '''
        <!doctype html>
            <title>Upload new File</title>
            <strong style="color:red;">Selecione as paginas</strong>
        <form method=post enctype=multipart/form-data style="margin-top:1rem">
            <input type=file name=file>
            <input type=text name=pages>
            <input type=submit value=Upload>
        </form>'''
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            convert_file(filename,pages)
            file = filename.replace('.pdf',"")
            download_file(f'{file}.xlsx')
            
            return redirect(url_for('download_file', name=f"{file}.xlsx"))
    return '''
    <!doctype html>
    <div>
    <title>Upload new File</title>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=text name=pages>
      <input type=submit>
      
    </form>
    </div>
    '''
    
if __name__ == '__main__':
    app.run_server(
               debug=False, 
               dev_tools_ui=False,
               dev_tools_props_check=False,
               dev_tools_serve_dev_bundles=False) 