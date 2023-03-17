#!/usr/bin/env python
# coding: utf-8

# In[7]:


import pandas as pd
import os
import win32com.client as win32

# get the current working directory
cwd = os.getcwd()

# Attachments directory
attachments_folder = os.path.join(cwd, "attachments")
# set the name for the Excel file
recipients_list = "recipients_list.xlsx"
# set the full paths to the excel file
recipients_file_path = os.path.join(cwd, recipients_list)

# Open Excel file
excel_file = pd.ExcelFile(recipients_file_path)

# Load distribution_list worksheet into a pandas dataframe
df = pd.read_excel(excel_file, 'distribution_list')

# Load Email worksheet into a pandas dataframe
email_df = pd.read_excel(excel_file, 'Email')

# Loop through distribution list
for index, row in df.iterrows():
    
    # Load email body and subject from Email worksheet
    email_body = email_df.iloc[0][1]
    email_subject = email_df.iloc[0][0]

    # Replace {NAME} placeholder in the email body
    email_body = email_body.replace('{NAME}', row['Name'])
    
    # Replace {ATTACHMENT} placeholder in the email body if there is an attachment
    if not pd.isna(row['Attachment']):
        email_body = email_body.replace('{ATTACHMENT}', row['Attachment'])
    else:
        email_body = email_body.replace('{ATTACHMENT}', "")
    
    # Replace \n with <br>
    email_body = email_body.replace('\n', '<br>')

    # Create Outlook email object
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    # Set recipient, CC, subject, and body of email
    mail.To = row['Email Receiver']
    mail.Subject = email_subject
    mail.HTMLBody = email_body

    # Check if there is a CC receiver and add it to the email if there is
    if not pd.isna(row['Email CC']):
        mail.CC = row['Email CC']

    # Check if there is an attachment file and add it to the email
    if not pd.isna(row['Attachment']):
        attachment_path = os.path.join(attachments_folder, row['Attachment'])
        # Add it
        mail.Attachments.Add(attachment_path)
    else:
        pass
    
    # Display email using Outlook
    mail.Display()
    
# Close Excel file
excel_file.close()

# Print success message
print('All emails were successfully opened in Outlook!')

