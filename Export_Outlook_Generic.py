# -*- coding: utf-8 -*-
"""
Created on Fri Mar 21 09:00:29 2025

@author: eleni.georganta
"""

import win32com.client
import pandas as pd

#The email adresses that you want to exclude from your query
excluded_addresses = ["examplemail@mail.com", "examplemail2@mail.com"]

#Looking into each subfolder
def extract_emails_from_folder(folder, excluded_addresses=None, start_date=None, end_date=None):
    email_data = []
    #print(f"Processing Folder: {folder.Name} - Total Items: {len(folder.Items)}") #debug line
    # Extract emails from the current folder
    messages = folder.Items
    messages.Sort("[CreationTime]", True)

    if start_date and end_date:
        filter_query = f"[CreationTime] >= '{start_date}' AND [CreationTime] <= '{end_date}'"
        messages = messages.Restrict(filter_query)
    
    for message in messages:
        try:
            #print(f"Sender Email: {message.SenderEmailAddress}") #debug line
            if message.Class != 43:  # it'a not a mail item
                #print(f"Skipped: Non-Mail Item in {folder.Name}") #debug line
                continue
            if not hasattr(message, 'ReceivedTime'):
                #print(f"Skipped: Missing ReceivedTime - {message.Subject}") #debug line
                continue
            sender_email = message.SenderEmailAddress.strip().lower() if message.SenderEmailAddress else None
            if excluded_addresses and sender_email in [addr.lower() for addr in excluded_addresses]:
                #print(f"Skipped: Excluded Address - {sender_email}") #debug line
                continue  
            email_info = {
                "Folder": folder.Name,
                "Subject": message.Subject if hasattr(message, 'Subject') else "No Subject",
                "Sender": message.SenderName if hasattr(message, 'SenderName') else "Unknown Sender",
                "ReceivedTime": message.ReceivedTime.strftime("%d-%m-%Y"),
                "Body": message.Body[:1500] if hasattr(message, 'Body') else "No Body"
            }
            email_data.append(email_info)
        except Exception as e:
            print(f"Error processing message: {e}") 
    
    #Subfolders
    for subfolder in folder.Folders:
        #print(f"Recursing into Subfolder: {subfolder.Name}") #debug line
        email_data.extend(extract_emails_from_folder(subfolder, excluded_addresses, start_date, end_date))
    
    return email_data

#Retriving mails with the help of "extract_emails_from_folder" def
def extract_outlook_emails(mailbox_name, excluded_addresses=None, start_date=None, end_date=None):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        mailbox = None
        for folder in outlook.Folders:
            if folder.Name == mailbox_name:
                mailbox = folder
                break
        
        if not mailbox:
            raise ValueError(f"Mailbox '{mailbox_name}' not found.")
        
        email_data = extract_emails_from_folder(mailbox, excluded_addresses, start_date, end_date)
        return email_data

    except Exception as e:
        print(f"An error occurred: {e}")
        return []

#Save mails to excel
def save_emails_to_excel(emails, file_path="Outlook_Emails.xlsx"):
    try:
        df = pd.DataFrame(emails)
        pivot = df.groupby("Folder").size().reset_index(name="Email Count")
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Emails breakdown", index=False)
            pivot.to_excel(writer, sheet_name="Pivot", index=False)
        print(f"Emails successfully saved to {file_path}")
    except Exception as e:
        print(f"An error occurred while saving to Excel: {e}")

#Setting up mailbox and timeframe 
if __name__ == "__main__":
    mailbox_name = "Mail Box Name" #update accordingly
    start_date = "1-1-2025"  #(DD-MM-YYYY)
    end_date = "31-1-2025"  #(DD-MM-YYYY)
    output_file = f"Outlook_Emails {start_date}_{end_date}.xlsx" 

    # Extract emails 
    emails = extract_outlook_emails(mailbox_name, excluded_addresses, start_date, end_date)
    
    #Confirmation or error message
    if emails:
        print(f"Succesfully fetched {len(emails)} emails from '{mailbox_name}' and its subfolders.")
        save_emails_to_excel(emails, output_file)
    else:
        print("No emails found or an error occurred.")
