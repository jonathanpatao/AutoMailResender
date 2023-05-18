import win32com.client as win32
from pdfreader import SimplePDFViewer
import os
import numpy as np
import re


class MailHandler:
    def __init__(self):
        self.outlook = win32.Dispatch('outlook.application').GetNamespace('MAPI')

        self.automation_folder = self.outlook.Folders[0].Folders['Automation']

        self.valid_file_types = [
            '.pdf',
            '.docx',
            '.doc'
        ]
        self.template_message_path = ''
        pass

    def get_unprocessed_messages(self):
        messages = self.automation_folder.Items
        return messages

    def delete_unprocessed_messages(self):
        return
        num_msgs = self.automation_folder.Items.Count
        for _ in range(num_msgs):
            self.automation_folder.Items.GetFirst().Delete()


    def is_valid_message(self, message):
        if message.Attachments.Count == 0:
            return False

        for attachment in message.Attachments:
            file_name = attachment.FileName
            file_type = os.path.splitext(file_name)[1]
            if file_type in self.valid_file_types:
                return True

        return False

    def save_attachments(self, message_idx, message, path):
        # init params:
        file_path_list = []
        real_file_name_list = []

        for i in range(message.Attachments.Count):
            attachment = message.Attachments[i]

            file_name = attachment.FileName
            file_type = os.path.splitext(file_name)[1]
            if file_type in self.valid_file_types:
                save_file_path = os.path.join(path, f'file_{message_idx}_{i}{file_type}')
                attachment.SaveAsFile(save_file_path)
                file_path_list.append(save_file_path)
                real_file_name_list.append(file_name)

        return file_path_list, real_file_name_list


    def send_new_mail(self, mail_list: list):
        if mail_list == []:
            pass
        to = '; '.join(mail_list)
        # print(to)
        pass

    @staticmethod
    def delete_message(message):
        message.Delete()



