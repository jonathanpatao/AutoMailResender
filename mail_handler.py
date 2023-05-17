import win32com.client as win32
from pdfreader import SimplePDFViewer
import os
import numpy as np
import re


class Outlook:
    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application').GetNamespace('MAPI')

        self.automation_folder = self.outlook.Folders[0].Folders['Automation']

        self.valid_file_types = [
            '.pdf',
            '.docx'
        ]
        self.template_message_path = ''
        pass

    def get_unprocessed_messages(self):
        return self.automation_folder.Items


    def is_valid_message(self, message):
        if message.Attachments.Count == 0:
            return False

        for attachment in message.Attachments:
            file_name = attachment.FileName
            file_type = os.path.splitext(file_name)[1]
            if file_type in self.valid_file_types:
                return True

        return False

    def save_attachments(self, message, path):
        # init params:
        file_path_list = []

        for attachment in message.Attachments:
            file_name = attachment.FileName
            file_type = os.path.splitext(file_name)[1]
            if file_type in self.valid_file_types:
                save_file_path = os.path.join(path, file_name)
                attachment.SaveAsFile(save_file_path)
                file_path_list.append(save_file_path)

        return file_path_list


    def send_new_mail(self, to):
        pass

    @staticmethod
    def delete_mail(message):
        message.Delete()



class MailHandler:
    def __init__(self, folder_name):
        self.outlook = win32.Dispatch('outlook.application').GetNamespace('MAPI')
        main_folder = self.outlook.Folders.GetFirst()
        self.folder = main_folder.Folders.Item(folder_name)

        self.email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        self.valid_extentions = ('.pdf', '.docx')
        self.temp_path = None

    def get_mail_from_attachment(self, attachment):
        # save file:
        with tempfile.TemporaryDirectory() as path:
            self.temp_path = path
            file_path = os.path.join(self.temp_path, attachment.FileName)
            attachment.SaveAsFile(file_path)

            file_type = os.path.splitext(file_path)[1]
            if file_type == '.pdf':
                mail = self.get_mail_from_pdf(file_path)
            elif file_type == '.docx':
                pass

            # delete te file:

    def get_mail_from_pdf(self, file_path):

        # Convert PDF to images
        images = convert_from_path(file_path, output_folder=self.temp_path)

        mail_list = []
        for i, image in enumerate(images):
            # Perform OCR on the image using pytesseract
            text = pytesseract.image_to_string(image)
            email_matches = re.findall(self.email_regex, text)
            # Print the extracted text
            mail_list += email_matches

    def run(self):
        unread_messages = self.folder.Items.Restrict("[Unread]=True")
        for message in unread_messages:
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    attachment_ext = os.path.splitext(attachment.FileName)[1]
                    if attachment_ext in self.valid_extentions:
                        mail = self.get_mail_from_attachment(attachment)


def read_from_pdf(filename):
    pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'

    with open(filename, "rb") as fd:
        viewer = SimplePDFViewer(fd)
        mail_content = []
        for canvas in viewer:
            text_content = canvas.text_content.split('\n')
            mail_content += [re.findall(pattern, s) for s in text_content if len(re.findall(pattern, s)) != 0]

    return mail_content


import re
import PyPDF2

def extract_emails_from_pdf(pdf_path):
    # Open the PDF file
    with open(pdf_path, 'rb') as f:
        pdf_reader = PyPDF2.PdfReader(f)
        # Loop through each page of the PDF
        for page in pdf_reader.pages:
            # Extract text from the page
            text = page.extract_text()
            # Use regular expressions to find email addresses in the text
            email_regex = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
            email_matches = re.findall(email_regex, text)
            # Return the email addresses found
            if email_matches:
                return email_matches

    # If no email addresses were found, return None
    return None


a = r'C:\Users\User\Downloads\Drushim_216215.pdf'
b = r'C:\Users\User\Downloads\Tehila Berdugo 15.05.23 CV.pdf'
c = r'C:\Users\User\Downloads\CV Example.pdf'
# x = extract_emails_from_pdf(c)
# x = extract_emails_from_pdf(a)
# x = extract_emails_from_pdf(b)

# from pdf2image import convert_from_path
#
# # Replace 'C:/path/to/your/file.pdf' with the actual path to your PDF file
# pages = convert_from_path(a)
#
# # The above function returns a list of PIL.Image objects.
# # You can then loop over the pages and save them to disk:
# for i, page in enumerate(pages):
#     page.save(f'page_{i}.jpg', 'JPEG')
#
# import os
# import tempfile
# from pdf2image import convert_from_path
# import pytesseract
#
# # Path to the PDF file
# pdf_path = c
# output_path = r'C:\temp\mail_handler'
#
# # Convert PDF to images
# with tempfile.TemporaryDirectory() as path:
#
#     images = convert_from_path(pdf_path, output_folder=path)
#
#     # Loop through each image and perform OCR
#     for i, image in enumerate(images):
#         # Save the image to disk (optional)
#         # image.save(os.path.join(path, f'{i}.jpg'))
#
#         # Perform OCR on the image using pytesseract
#         text = pytesseract.image_to_string(image)
#         email_matches = re.findall(email_regex, text)
#         # Print the extracted text
#         print('; '.join(email_matches))

a = Outlook()
pass

