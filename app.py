from flask import Flask, request, send_file, render_template
import os
import zipfile
import PyPDF2
import docx2txt
import pandas as pd
import re
import shutil
import io
from spire.doc import *
from spire.doc.common import *
import warnings

warnings.filterwarnings("ignore")
app = Flask(__name__)


# Function to extract text from PDF
def extract_text_from_pdf(file_path):
  """Extract text from a PDF file using PyPDF2."""
  text = ''
  try:
    with open(file_path, 'rb') as file:
      pdf_reader = PyPDF2.PdfReader(file)
      num_pages = len(pdf_reader.pages)
      for page_num in range(num_pages):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
  except Exception as e:
    print(f"Error extracting text from PDF file '{file_path}': {e}")
  return text


# Function to extract text from DOC/DOCX
def extract_text_from_doc_or_docx(file_path):
  """Extract text from a DOC or DOCX file using docx2txt."""
  try:
    text = docx2txt.process(file_path)
  except Exception as e:
    print(f"Error extracting text from DOC/DOCX file '{file_path}': {e}")
    text = ''
  return text


def extract_text_from_doc(file_path):
  # Create a Document object
  document = Document()
  # Load a Word document
  document.LoadFromFile(file_path)
  # Extract the text of the document
  document_text = document.GetText()
  document_text = document_text.replace(
      "Evaluation Warning: The document was created with Spire.Doc for Python.",
      "")
  document.Close()
  return document_text


# Function to extract email and contact information from text

def extract_email_and_contact(text):
    """Extract email and contact information from text."""
    
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    contact_pattern = r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'
  
    # Find the first occurrence of email and contact
    email = re.search(email_pattern, text)
    contact = re.search(contact_pattern, text)

    # Extract email and contact if found, otherwise return empty strings
    email = email.group() if email else ''
    contact = contact.group() if contact else ''

    return email, contact




# Function to process files in a directory
def process_directory(directory, data):
  """Recursively process all files and directories in a given directory."""
  for item in os.listdir(directory):
    item_path = os.path.join(directory, item)

    if os.path.isfile(item_path):
      # Determine file format and extract text accordingly
      if item.endswith('.pdf'):
        text = extract_text_from_pdf(item_path)
      elif item.endswith('.docx'):
        text = extract_text_from_doc_or_docx(item_path)
      elif item.endswith('.doc'):
        text = extract_text_from_doc(item_path)
      else:
        print(f"Unsupported file format for '{item}'")
        continue

      # Extract email and contact information
      email, contact = extract_email_and_contact(text)

      # Add data to list
      data.append({
          'File Name': item,
          'Email': email,
          'Contact': contact,
          'Text': text
      })

      # Delete the file after processing
      os.remove(item_path)
      print(f"Deleted '{item_path}'.")

    elif os.path.isdir(item_path):
      # Recursively process directories
      process_directory(item_path, data)
      # After processing, delete the empty directory
      shutil.rmtree(item_path)
      print(f"Deleted directory '{item_path}'.")


# Function to handle ZIP file processing and create Excel file
def handle_zip_file(zip_file_path, extract_to_path):
  # Ensure the extract_to_path directory exists
  os.makedirs(extract_to_path, exist_ok=True)

  # Extract the ZIP file
  with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
    zip_ref.extractall(extract_to_path)
    print(f"Extracted contents from '{zip_file_path}' to '{extract_to_path}'.")

  # List to store data
  data = []

  # Process the extraction path
  process_directory(extract_to_path, data)

  # Create a DataFrame from the data list
  df = pd.DataFrame(data)

  # Create an in-memory buffer for the Excel file
  excel_buffer = io.BytesIO()

  # Save the DataFrame to an Excel file in memory
  df.to_excel(excel_buffer, index=False)

  # Reset the buffer's position to the beginning
  excel_buffer.seek(0)

  # Return the in-memory Excel file
  return excel_buffer


# Route to render the upload form
@app.route('/')
def index():
  return render_template('index.html')


# Route to handle file upload and processing
@app.route('/upload', methods=['POST'])
def upload():
  # Check if the request contains a file
  if 'zip_file' not in request.files:
    return 'No file part in the request.', 400

  # Get the uploaded ZIP file
  zip_file = request.files['zip_file']

  # Check if a file was uploaded
  if zip_file.filename == '':
    return 'No file selected for uploading.', 400

  # Create a temporary directory to store the extracted files
  temp_extract_path = 'temp_extract'

  # Save the ZIP file temporarily
  zip_file_path = 'temp.zip'
  zip_file.save(zip_file_path)

  # Process the ZIP file and create an Excel file in memory
  excel_buffer = handle_zip_file(zip_file_path, temp_extract_path)

  # Clean up temporary files and directories
  os.remove(zip_file_path)
  shutil.rmtree(temp_extract_path)

  # Return the Excel file for download
  return send_file(
      excel_buffer,
      as_attachment=True,
      download_name='output.xlsx',
      mimetype=
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# Run the Flask app
if __name__ == '_main_':
  app.run(debug=True)
