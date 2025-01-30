import os
import sys
import win32com.client
from datetime import datetime

# Function to check and install dependencies
def check_install_dependencies():
    try:
        import win32com.client
    except ImportError:
        print("pywin32 not installed. Installing...")
        os.system("pip install pywin32")

# Function to extract attachments
def extract_attachments(subject_contains, attachment_name_contains, sender_address_contains, email_folder, save_location, save_in_individual_folders):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folder = outlook.Folders[email_folder]
    messages = folder.Items

    for message in messages:
        if (subject_contains.lower() in message.Subject.lower() and
            attachment_name_contains.lower() in message.Attachments.Item(1).FileName.lower() and
            sender_address_contains.lower() in message.SenderEmailAddress.lower()):
            
            for attachment in message.Attachments:
                save_path = save_location
                
                if save_in_individual_folders:
                    folder_name = f"{message.Subject}_{message.ReceivedTime.strftime('%Y-%m-%d')}"
                    save_path = os.path.join(save_location, folder_name)
                    if not os.path.exists(save_path):
                        os.makedirs(save_path)
                
                attachment.SaveAsFile(os.path.join(save_path, attachment.FileName))
                print(f"Attachment {attachment.FileName} saved to {save_path}")

# Main function
def main():
    check_install_dependencies()
    
    subject_contains = input("Enter keyword(s) contained in the email subject: ")
    attachment_name_contains = input("Enter keyword(s) contained in the attachment name: ")
    sender_address_contains = input("Enter keyword(s) contained in the sender's email address: ")
    email_folder = input("Enter the email folder where to search (e.g., 'Inbox'): ")
    save_location = input("Enter the location where the attachments should be saved: ")
    save_in_individual_folders = input("Save attachments in individual folders? (yes/no): ").lower() == 'yes'
    
    extract_attachments(subject_contains, attachment_name_contains, sender_address_contains, email_folder, save_location, save_in_individual_folders)

if __name__ == "__main__":
    main()
