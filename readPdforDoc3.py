import pyttsx3
from docx import Document
import fitz  # PyMuPDF
import os

def read_word_document(doc_path, start_page=1, speech_rate=150):
    # Initialize the text-to-speech engine
    engine = pyttsx3.init()
    
    # Set the speech rate
    engine.setProperty('rate', speech_rate)

    # Load the Word document
    if not os.path.isfile(doc_path):
        print(f"File path not found: {doc_path}")
        return

    try:
        doc = Document(doc_path)
    except Exception as e:
        print(f"Error loading document: {e}")
        return

    # Iterate through the paragraphs in the document starting from the specified page
    for i, paragraph in enumerate(doc.paragraphs, start=1):
        if i >= start_page:
            text = paragraph.text
            if text:  # Only read non-empty paragraphs
                engine.say(text)
       
    # Wait for the speech to finish
    engine.runAndWait()

def read_pdf_document(doc_path, start_page=1, speech_rate=150):
    # Initialize the text-to-speech engine
    engine = pyttsx3.init()
    
    # Set the speech rate
    engine.setProperty('rate', speech_rate)

    # Load the PDF document
    if not os.path.isfile(doc_path):
        print(f"File path not found: {doc_path}")
        return

    try:
        doc = fitz.open(doc_path)
    except Exception as e:
        print(f"Error loading document: {e}")
        return

    # Extract and read text from each page starting from the specified page
    for page_num in range(start_page - 1, len(doc)):
        try:
            page = doc.load_page(page_num)
            text = page.get_text()
            if text:  # Only read non-empty text
                engine.say(text)
        except Exception as e:
            print(f"Error reading page {page_num}: {e}")

    # Close the PDF document
    doc.close()

    # Wait for the speech to finish
    engine.runAndWait()

def get_speech_rate():
    while True:
        try:
            return int(input("Please enter the speech rate (default is 150): ") or 150)
        except ValueError:
            print("Invalid input. Please enter a numeric value for the speech rate.")

def get_start_page():
    while True:
        try:
            start_page = int(input("Please enter the start page number (default is 1): ") or 1)
            if start_page > 0:
                return start_page
            else:
                print("Page number must be greater than 0.")
        except ValueError:
            print("Invalid input. Please enter a numeric value for the start page.")

def main():
    while True:
        # Display a menu for the user
        print("\n===========================================================")
        print("Welcome to the Word document or PDF reader!")
        print("============= Script written by Dagimbirhanu ===============")
        print("1. Read a Word document")
        print("2. Read a PDF document")
        print("3. Exit")
        choice = input("Please select an option (1, 2, or 3): ")

        if choice == '1':
            # Read Word document
            doc_path = input("Please enter the path to the Word document: ")
            start_page = get_start_page()
            speech_rate = get_speech_rate()
            read_word_document(doc_path, start_page, speech_rate)
        elif choice == '2':
            # Read PDF document
            doc_path = input("Please enter the path to the PDF document: ")
            start_page = get_start_page()
            speech_rate = get_speech_rate()
            read_pdf_document(doc_path, start_page, speech_rate)
        elif choice == '3':
            # Exit the program
            print("Exiting the program. Goodbye!")
            break
        else:
            print("Invalid choice. Please select either 1, 2, or 3.")

if __name__ == "__main__":
    main()
