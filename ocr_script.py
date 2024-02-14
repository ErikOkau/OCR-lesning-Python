import imaplib
import email
import os
import pytesseract
from PIL import Image
from PyPDF2 import PdfReader
import keyboard

# Email credentials and server configuration
EMAIL = 'ocrtesting2910@outlook.com'
PASSWORD = 'TestingOCR'
IMAP_SERVER = 'imap-mail.outlook.com'
OUTPUT_DIR = 'output'

# Path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

# Function to convert PDF to text using OCR
def pdf_to_text(pdf_path):
    try:
        pdf_reader = PdfReader(pdf_path)
        text = ''
        for page_num in range(len(pdf_reader.pages)):
            text += pdf_reader.pages[page_num].extract_text()
        return text
    except Exception as e:
        return str(e)

# Function to perform OCR on an image
def perform_ocr(image_path):
    try:
        with Image.open(image_path) as img:
            text = pytesseract.image_to_string(img)
            return text
    except Exception as e:
        return str(e)

# Function to process emails and convert PDF attachments
def process_emails():
    try:
        # Connect to the IMAP server
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL, PASSWORD)
        mail.select("inbox")

        # Search for all unseen emails
        status, messages = mail.search(None, 'UNSEEN')
        messages = messages[0].split()

        for mail_id in messages:
            _, msg_data = mail.fetch(mail_id, '(RFC822)')
            _, msg = msg_data[0]
            email_message = email.message_from_bytes(msg)

            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is not None:
                    filename = part.get_filename()
                    if filename.lower().endswith('.pdf'):
                        # Save the PDF attachment
                        pdf_path = os.path.join(OUTPUT_DIR, filename)
                        with open(pdf_path, 'wb') as pdf_file:
                            pdf_file.write(part.get_payload(decode=True))

                        # Convert PDF to text using OCR
                        extracted_text = pdf_to_text(pdf_path)

                        # Save extracted text to a text file
                        text_file_path = os.path.join(OUTPUT_DIR, f'{os.path.splitext(filename)[0]}.txt')
                        with open(text_file_path, 'w', encoding='utf-8') as text_file:
                            text_file.write(extracted_text)

                        print(f'Successfully converted {filename} to text: {text_file_path}')

        mail.close()
        mail.logout()

    except Exception as e:
        print(f'Error: {str(e)}')

# Create the output directory if it doesn't exist
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Call the function to process emails and convert PDF attachments
process_emails()



# Function to listen for keyboard shortcut and process emails
def listen_for_shortcut():
    # Define the keyboard shortcut (for example, Ctrl + Shift + O)
    shortcut = 'ctrl+shift+o'

    while True:
        if keyboard.is_pressed(shortcut):
            print("Shortcut pressed. Processing emails...")
            process_emails()
            print("Email processing complete.")
            # Optional: You can add a delay to prevent multiple rapid executions if the shortcut is held down
            keyboard.wait('esc', suppress=True, trigger_on_release=True)

# Call the function to listen for the shortcut and process emails
listen_for_shortcut()
