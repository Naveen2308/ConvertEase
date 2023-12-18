from flask import Flask, render_template, request, redirect, send_from_directory
from pdf2docx import Converter
import os

def delete_files():
    static_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')
    try:
        for filename in os.listdir(static_folder):
            file_path = os.path.join(static_folder, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted file: {file_path}")
    except Exception as e:
        print(f"Error deleting files: {e}")

app = Flask(__name__, static_folder='static')
app.config['STATIC_FOLDER'] = 'static'
@app.route('/index.html', methods=['GET', 'POST'])
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'GET' and request.path == '/index.html' or request.method == 'GET' and request.path == '/':
        delete_files()

    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            return redirect(request.url)

        if file and file.filename.endswith('.pdf'):
            pdffile = file.filename
            wordfile = pdffile.replace('.pdf', '.docx')

            # Specify the path to the static folder
            static_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')

            # Save the PDF file in the static folder
            pdf_path = os.path.join(static_folder, pdffile)
            file.save(pdf_path)

            cv = Converter(pdf_path)

            # Specify the path to save the Word file in the static folder
            word_path = os.path.join(static_folder, wordfile)
            
            # Convert the PDF to Word and save it in the static folder
            cv.convert(word_path)
            cv.close()

            return render_template('result.html', wordfile=wordfile)
        


    return render_template('index.html')
@app.route('/about.html', methods=['GET', 'POST'])
def about():
    return render_template('about.html')
@app.route('/downloads/<filename>')
def download(filename):
    # Serve the file from the static folder
    return send_from_directory('static', filename, as_attachment=True)

if __name__ == '__main__':
    # Create 'static' folder if it doesn't exist
    static_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'static')
    os.makedirs(static_folder, exist_ok=True)


    app.run(debug=True)

    for filename in os.listdir(static_folder):
        file_path = os.path.join(static_folder, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
        except Exception as e:
            print(f"Error deleting file {file_path}: {e}")
