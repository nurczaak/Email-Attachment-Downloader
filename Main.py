import imaplib
import email
import os
from datetime import datetime, timedelta
from email.header import decode_header
import tkinter as tk
from tkinter import messagebox, filedialog


def fetch_attachments():
    # Email configuration
    username = entry_username.get()
    password = entry_password.get()
    sender_email = entry_sender_email.get()
    imap_server = 'imap.wp.pl'
    save_folder = entry_directory.get()  # Use the directory entered by the user
    if not os.path.exists(save_folder):
        os.makedirs(save_folder)  # Create the directory if it doesn't exist

    # Get start and end dates from entry fields
    start_date_str = entry_start_date.get()
    end_date_str = entry_end_date.get()

    # Convert date strings to datetime objects
    start_date = datetime.strptime(start_date_str, "%Y-%m-%d")
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d")

    # Connect to the IMAP server
    mail = imaplib.IMAP4_SSL(imap_server, 993)
    try:
        mail.login(username, password)
        mail.select('INBOX')

        # Search for emails from the specific sender within the date range
        date_criteria = f'(SINCE "{start_date.strftime("%d-%b-%Y")}") (BEFORE "{(end_date + timedelta(days=1)).strftime("%d-%b-%Y")}")'
        typ, msgnums = mail.search(None, f'(FROM "{sender_email}")', date_criteria)

        # Fetch and save attachments from qualifying emails
        for num in msgnums[0].split():
            typ, msg_data = mail.fetch(num, '(RFC822)')
            raw_email = msg_data[0][1]
            email_message = email.message_from_bytes(raw_email)

            # Check for attachments and save them
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                # Get the attachment filename and decode it properly
                filename = part.get_filename()
                filename_decoded = decode_header(filename)[0][0]
                filename_charset = decode_header(filename)[0][1]
                if filename_charset:
                    filename_decoded = filename_decoded.decode(filename_charset)

                # Save attachments to the specified folder
                if filename_decoded:
                    filepath = os.path.join(save_folder, filename_decoded)
                    open(filepath, 'wb').write(part.get_payload(decode=True))
                    result_text.insert(tk.END, f"Attachment '{filename_decoded}' saved in '{save_folder}'.\n")

        mail.logout()
        messagebox.showinfo("Success", "Attachments fetched and saved successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")
        mail.logout()



root = tk.Tk()
root.title("Email Attachment Fetcher")

label_username = tk.Label(root, text="Enter Email:")
label_username.pack()
entry_username = tk.Entry(root)
entry_username.pack()


label_password = tk.Label(root, text="Enter Password:")
label_password.pack()
entry_password = tk.Entry(root, show="*")
entry_password.pack()


label_sender_email = tk.Label(root, text="Enter Sender Email:")
label_sender_email.pack()
entry_sender_email = tk.Entry(root)
entry_sender_email.pack()

label_directory = tk.Label(root, text="Enter Directory to Save Attachments:")
label_directory.pack()
entry_directory = tk.Entry(root)
entry_directory.pack()

label_start_date = tk.Label(root, text="Enter Start Date (YYYY-MM-DD):")
label_start_date.pack()
entry_start_date = tk.Entry(root)
entry_start_date.pack()


label_end_date = tk.Label(root, text="Enter End Date (YYYY-MM-DD):")
label_end_date.pack()
entry_end_date = tk.Entry(root)
entry_end_date.pack()



def browse_directory():
    selected_directory = filedialog.askdirectory()
    entry_directory.delete(0, tk.END)
    entry_directory.insert(tk.END, selected_directory)


browse_button = tk.Button(root, text="Browse", command=browse_directory)
browse_button.pack()


fetch_button = tk.Button(root, text="Fetch Attachments", command=fetch_attachments)
fetch_button.pack()


result_text = tk.Text(root, height=10, width=50)
result_text.pack()

root.mainloop()
