import os
import time
import tempfile
import numpy as np
from tqdm import tqdm

import win32com.client as win32
import pandas as pd
from datetime import datetime, timedelta

from mail_handler import MailHandler
from file_parser import PdfParser, WordParser, filter_list

ACCOUNT_MAIL = 'jonathan.patao1995@gmail.com'


def statistics():
    # Init Variables
    exception_flag = 0
    run_messages = 0
    parsed_messages = 0
    valid_messages = 0
    real_name = None

    total_time = time.time()

    # Start Script
    mail = MailHandler(ACCOUNT_MAIL)

    # Get Unprocessed messages:
    messages = mail.get_unprocessed_messages()
    nof_messages = len(messages)
    time_list = []

    # Open Temporary folder to save the files:
    with tempfile.TemporaryDirectory() as temp_path:
        try:
            # Loop through the messages.
            # Because we delete the messages at the end of the loop, use while instead of for.
            # while len(messages) > 0:
            print(f"There is {nof_messages} messages.")
            if nof_messages == 0:
                return

            for message_idx in range(nof_messages):
                message = messages[message_idx]  # messages[message_idx]
                # Work only on valid messages
                if mail.is_valid_message(message):
                    valid_messages += 1
                    start_time = time.time()

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
                    if len(filtered_mail_list) > 0:
                        parsed_messages += 1
                end_time = time.time()
                time_list.append(end_time - start_time)
                run_messages += 1
            total_time = time.time() - total_time

        except Exception as e:
            exception_flag = 1
            print(e)
            print(f"Stoped at file: {real_name}")

    if exception_flag:
        print(f"ERROR: the script stops after {run_messages} in {nof_messages} messages.")
    else:
        # calculate statistics:
        pst_valid = round(100 * (valid_messages / nof_messages), 1)
        pst_parsed = round(100 * (parsed_messages / nof_messages), 1)

        mean_time = round(float(np.mean(time_list)), 0)
        std_time = round(float(np.std(time_list)), 0)
        max_time = round(float(np.max(time_list)), 0)
        min_time = round(float(np.min(time_list)), 0)
        med_time = round(float(np.median(time_list)), 0)

        print(f"Done Proess {nof_messages} messages.\n")
        print(' ')
        print('Time Analisys:')
        print(f'Total Time: {time.strftime("%H:%M:%S", time.gmtime(total_time))}')
        print(f'processed message duration: average of {mean_time} seconds, and std of {std_time} seconds.')
        print(f'processed message duration: max of {max_time} seconds, and min of  {min_time} seconds.')
        print(f'processed message duration: median of {med_time} seconds.\n')
        print(' ')
        print('Process Analysis:')
        print(f"Valid Messages: total of {valid_messages} messages, and {pst_valid}% from {nof_messages} messages.")
        print(f"Valid Messages: total of {parsed_messages} messages, and {pst_parsed}% from {nof_messages} messages.")


def get_messages_statistics():
    # Connect to Outlook
    outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 represents the Inbox folder

    # Get the start and end dates for the analysis
    end_date = datetime.today().date()
    start_date = end_date - timedelta(days=90)

    # Create a DataFrame to store the results
    data = {'Date': pd.date_range(start=start_date, end=end_date)}
    df = pd.DataFrame(data)
    df['Messages'] = 0
    dates_list = np.array([df['Date'][i].date() for i in range(len(df))])

    # Restrict items based on received date and attachments
    restriction = (
        "[ReceivedTime] >= '{}' AND "
        "[ReceivedTime] <= '{}'"
    ).format(
        start_date.strftime("%m/%d/%Y"),
        end_date.strftime("%m/%d/%Y")
    )
    restricted_items = inbox.Items.Restrict(restriction)
    messages_list = [msg for msg in restricted_items]

    # Retrieve messages and count per day
    for message in messages_list:
        if message.Attachments.Count>0:
            valid = True
            for attachment in message.Attachments:
                file_type = os.path.splitext(attachment.FileName)[1]
                if file_type in ['.pdf', '.doc', '.docx']:
                    valid = True
            if valid:
                index = df.loc[dates_list == message.ReceivedTime.date()].index.to_list()[0]
                df.at[index, 'Messages'] += 1

    # Print the analysis results
    print(df)

    # Save the results to a CSV file
    df.to_csv('message_analysis.csv', index=False)



if __name__ == '__main__':
    try:
        get_messages_statistics()
    except Exception as e:
        print(e)
