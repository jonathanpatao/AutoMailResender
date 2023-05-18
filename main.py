from mail_handler import MailHandler
from file_parser import WordParser, PdfParser, filter_list
import tempfile
import os

def main():
    # Init Variables
    mail = MailHandler()

    # Get Unprocessed messages:
    messages = mail.get_unprocessed_messages()

    # Open Temporary folder to save the files:
    with tempfile.TemporaryDirectory() as temp_path:

        # Loop through the messages.
        # Because we delete the messages at the end of the loop, use while instead of for.
        # while len(messages) > 0:
        k=1
        for message_idx in range(messages.Count):
            message = messages[message_idx]
            # Work only on valid messages
            if mail.is_valid_message(message):
                # get the
                attachments_list, real_filename_list = mail.save_attachments(message_idx, message, temp_path)
                mail_list = []
                for i in range(len(attachments_list)):
                    file_path = attachments_list[i]
                    real_name = real_filename_list[i]

                    # Extruct mails only from .pdf or .docx
                    file_type = os.path.splitext(file_path)[1]
                    if file_type == '.pdf':
                        parser = PdfParser(file_path, temp_path)
                    elif file_type in ['.docx', '.doc']:
                        parser = WordParser(file_path)
                    else:
                        continue

                    mail_list += parser.extruct_mail()

                # Filterring the list to get only unique values
                filtered_mail_list = filter_list(mail_list)

                # Sending a template mail
                mail.send_new_mail(filtered_mail_list)
                print(f"number {k}: is Valid   the mail is: {filtered_mail_list[0] if len(filtered_mail_list) > 0 else ''}")
                pass
            else:
                print(f"number {k}: is not Valid.")
            k = k+1


    # Delete all the message after Done Process
    mail.delete_unprocessed_messages()
    pass




if __name__ == '__main__':
    main()