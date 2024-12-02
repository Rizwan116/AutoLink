from flask import Flask, request, render_template, send_file
import fitz  # PyMuPDF
import os
import pandas as pd

app = Flask(__name__)
@app.route('/')
def home():
    return render_template('index.html')

@app.route('/new_pdf_linking')
def new_pdf_linking():
    return render_template('new_pdf_linking.html')

@app.route('/about')
def about():
    return render_template('about.html')

@app.route('/contact')
def contact():
    return render_template('contact.html')

if __name__ == '__main__':
    app.run(debug=True)

app.config['UPLOAD_FOLDER'] = 'uploads'

def read_links_from_excel(file_path):
    links = []
    # Read the Excel file
    df = pd.read_excel(file_path)
    # Assuming the first column is for page numbers and the second for URLs
    for index, row in df.iterrows():
        try:
            page_num = int(row[0]) - 1  # Convert to zero-index
            url = row[1]
            links.append((page_num, url))
        except (ValueError, IndexError):
            print(f"Invalid data in row {index + 1}: {row}")
    return links

def replace_pdf_links(pdf_path, excel_file_path, output_pdf_path):
    links = read_links_from_excel(excel_file_path)
    pdf_document = fitz.open(pdf_path)
    link_index = 0

    for page_num in range(pdf_document.page_count):
        page = pdf_document[page_num]
        for annot in page.annots():
            if annot.type[0] == 1:
                page.delete_annot(annot)
        link_boxes = page.get_links()
        for link_box in link_boxes:
            if link_index < len(links):
                _, url = links[link_index]
                rect = link_box['from']
                page.insert_link({
                    "kind": fitz.LINK_URI,
                    "from": rect,
                    "uri": url
                })
                link_index += 1
            if link_index >= len(links):
                break

    pdf_document.save(output_pdf_path)
    pdf_document.close()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        pdf_file = request.files['pdf_file']
        excel_file = request.files['excel_file']
        
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
        output_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], "updated_" + pdf_file.filename)
        
        pdf_file.save(pdf_path)
        excel_file.save(excel_file_path)
        
        replace_pdf_links(pdf_path, excel_file_path, output_pdf_path)
        
        return send_file(output_pdf_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)


    
