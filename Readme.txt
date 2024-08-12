PDF and Word Document Reader through Voice

Overview
This Python project allows users to read PDF and Word documents aloud using a text-to-speech engine. Users can specify the page number from which they want to start reading, and the script will handle the rest.

Features
Supports both PDF and Word documents: The script can read from .pdf and .docx files.
Customizable speech rate: Users can set the speed at which the text is read aloud.
Start reading from a specific page: Users can choose the page number from which the reading should begin.
User-friendly interface: A simple menu guides users to select options easily.
Requirements
Python 3.x
PyMuPDF (fitz)
python-docx
pyttsx3


Installation
Clone the Repository:
git clone https://github.com/Dagimbirhanu-B/read-pdf-and-document-through-voice.git

cd read-pdf-and-document-through-voice

Install Dependencies:
You can install the required Python libraries using pip:

pip install -r requirements.txt

Alternatively, you can install them individually:
pip install PyMuPDF python-docx pyttsx3

Usage
Run the Script:

python readPdforDoc3.py


Choose the Document Type:

Option 1: Read a Word document.
Option 2: Read a PDF document.
Enter the Document Path:
Provide the path to the document you want to read.

Set the Speech Rate (Optional):
Enter the desired speech rate (default is 150).

Specify the Starting Page:
Enter the page number from which you want the reading to start.

Listen to the Document:
The script will read the document aloud starting from the specified page.


Example 

Welcome to the Word document or PDF reader!
============= Script written by Dagimbirhanu ===============
1. Read a Word document
2. Read a PDF document
3. Exit
Please select an option (1, 2, or 3): 2
Please enter the path to the PDF document: sample.pdf
Please enter the speech rate (default is 150): 150
Please enter the page number to start reading from: 3
