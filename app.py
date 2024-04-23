import os
from flask import Flask, flash, request, redirect,url_for,jsonify,make_response
from werkzeug.utils import secure_filename
from flask import send_from_directory
from spire.pdf.common import *
from spire.pdf import *
import fitz
from flask_cors import CORS, cross_origin

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}

app = Flask(__name__)
app.secret_key = b'_5#y2L"F4Q8z\n\xec]/'

CORS(app, resources={r"*": {"origins": "*"}})
app.config['CORS_HEADERS'] = 'Content-Type'

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


@cross_origin()
@app.route('/uploads/<name>')
def download_file(name):
    response = make_response(send_from_directory(app.config["UPLOAD_FOLDER"], name))
    response.headers['Content-Disposition'] = f'attachment; filename={name}'
    response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'
    return send_from_directory(app.config["UPLOAD_FOLDER"], name)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
           
@cross_origin()
@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        att = 'Gerando arquivo...'
        
        # check if the post request has the file part
        if 'file[]' not in request.files:
            flash('No file part')
            print('No file part')
            return jsonify("Arquivo não encontrado"), 404
        print(request.form)
        file = request.files['file[]']
        pages = request.form['pages']
        print(pages)
        print(file.filename)
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if file.filename == '':
            #flash('No selected file'
            return jsonify("Arquivo não encontrado"), 404
        if pages == '':
            
            #flash('No selected file')
            return jsonify("Paginas não selecionadas"), 400
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            convert_file(filename,pages)
            file = filename.replace('.pdf',"")
            download_file(f'{file}.xlsx')
            
            try:
                
                return redirect(url_for('download_file', name=f"{file}.xlsx")) 
            except:
                return jsonify("Houve um erro ao tentar enviar o arquivo")
    
    
if __name__ == '__main__':
    app.run_server(
               debug=False, 
               dev_tools_ui=False,
               dev_tools_props_check=False,
               dev_tools_serve_dev_bundles=False) 