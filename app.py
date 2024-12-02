from flask import Flask, request, render_template, send_file
import fitz  # PyMuPDF
import os
import pandas as pd

app = Flask(__name__)

# Folder to store uploaded files
app.config['UPLOAD_FOLDER'] = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Home route
@app.route('/')
def home():
    return render_template('index.html')

# Route to handle new PDF linking
@app.route('/new_pdf_linking')
def new_pdf_linking():
    return render_template('new_pdf_linking.html')

# About page route
@app.route('/about')
def about():
    return render_template('about.html')

# Contact page route
@app.route('/contact')
def contact():
    return render_template('contact.html')

# Function to read links from Excel file
def read_links_from_excel(file_path):
    links = []
    df = pd.read_excel(file_path)
    for index, row in df.iterrows():
        try:
            page_num = int(row[0]) - 1  # Convert to zero-indexed page number
            url = row[1]
            links.append((page_num, url))
        except (ValueError, IndexError):
            print(f"Invalid data in row {index + 1}: {row}")
    return links

def replace_pdf_links(pdf_path, excel_file_path, output_pdf_path):
    links = read_links_from_excel(excel_file_path)  # Get the new links from the Excel file
    pdf_document = fitz.open(pdf_path)  # Open the PDF document

    link_index = 0  # Keep track of which new link we are adding
    for page_num in range(pdf_document.page_count):  # Iterate over each page of the PDF
        page = pdf_document[page_num]  # Get the page object

        # Remove all existing links first
        current_links = page.get_links()
        for current_link in current_links:
            # print(f"Removing link on page {page_num + 1}: {current_link.get('uri', 'Unknown')}")
            page.delete_link(current_link)  # Delete the specific link

        # Add new links from the Excel file
        for i in range(min(len(current_links), len(links))):
            # Get the new URL from the Excel file
            _, new_url = links[link_index]
            
            # Get the existing link's rectangle (location)
            rect = current_links[i]['from']
            
            print(f"Adding new link on page {page_num + 1}: {new_url}")
            
            # Insert the new link at the same location
            page.insert_link({
                "kind": fitz.LINK_URI,  # Type of link is URI
                "from": rect,  # Location of the link
                "uri": new_url  # The new URL to insert
            })

            link_index += 1  # Move to the next link in the list
            
            if link_index >= len(links):  # Stop if we've added all the new links
                break

    # Save the updated PDF
    pdf_document.save(output_pdf_path)
    pdf_document.close() 
      
# Route to handle file upload and PDF link replacement
@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Get the uploaded PDF and Excel files
        pdf_file = request.files['pdf_file']
        excel_file = request.files['excel_file']

        # Save the uploaded files
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_file.filename)
        excel_file_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_file.filename)
        output_pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], "updated_" + pdf_file.filename)

        # Save the files to the server
        pdf_file.save(pdf_path)
        excel_file.save(excel_file_path)

        # Replace the links in the PDF
        replace_pdf_links(pdf_path, excel_file_path, output_pdf_path)

        # Return the updated PDF file for download
        return send_file(output_pdf_path, as_attachment=True)

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
