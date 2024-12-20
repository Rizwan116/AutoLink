# AutoLink Upload Files

AutoLink is a web application that allows users to upload PDF and Excel files. It processes the uploaded files and updates links in the PDF based on the data provided in the Excel file.

#Live Link 
https://autolink.onrender.com/

## Features

- User-friendly interface for uploading files.
- Supports PDF and Excel file formats.
- Processes links from the uploaded Excel file to update the PDF.

- ## Excel Format Should be in This Format


| Page | URL     |       
| :-------- | :------- |
| `1`      | `https://www.example.com` |
| `2`      | `https://www.example2.com` |
| `3`      | `https://www.example3.com` |
| `4`      | `https://www.example4.com` |


## Technologies Used

- **Frontend**: HTML, CSS (Bootstrap for styling)
- **Backend**: Python (Flask)
- **Libraries**:
  - `PyMuPDF`: For manipulating PDF files
  - `pandas`: For handling Excel files
  - `Flask`: For creating the web application

## Installation

### Prerequisites

Before you begin, ensure you have the following installed:

- Python 3.x
- pip (Python package manager)

### Clone the Repository

```bash
git clone https://github.com/yourusername/autolink-upload-files.git
cd autolink-upload-files
