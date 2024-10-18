import imaplib
import smtplib
import email
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import decode_header
# import json
import os
import csv
import datetime



def log_email(email_data, log_file='processed_emails.csv'):
  """Log the details of the email to a CSV file """
  file_exists = os.path.isfile(log_file)
  
  with open(log_file, 'a', newline='') as csvfile:
    fieldnames = ['sender', 'recipient', 'subject', 'attachment', 'received_date']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    
    # Write the header onl if the file doesn't exist yet
    if not file_exists:
      writer.writeheader()
      
    # Write the email data to the log
    writer.writerow(email_data)

def is_email_logged(subject, attachment, received_date, log_file='processed_emails.csv'):
  """Check if the email with the given subject, attachment received_date is already logged."""
  if not os.path.exists(log_file):
    return False # If log doesn't exist, nothing is logged yet
  
  with open(log_file, 'r') as csvfile:
    reader = csv.DictReader(csvfile)
    
    # Ensure the CSV file has the necessary columns
    if 'subject' not in reader.fieldnames or 'attachment' not in reader.fieldnames or 'received_date' not in reader.fieldnames:
      raise ValueError("CSV file does not contain the necessary 'subject', 'attachment' or 'received_date' columns.")
    
    for row in reader:
      if row['subject'] == subject and row['attachment'] == attachment and row['received_date'] == received_date:
        return True # Email is already logged
  
  return False

# Function to log in to the email server and search for emails
def search_emails(username, password, sender_email):
  # Connect to the email server
  mail = imaplib.IMAP4_SSL('YOUR IMAP DETAILS HERE')
  mail.login(username, password)
  
  # Select mailbox we want to search (e.g., "inbox")
  mail.select("inbox")
    
  # Select email sent from specific sender 
  status, messages = mail.search(None, '(FROM "{}")'.format(sender_email))
  
  email_ids = messages[0].split() # List of email IDs
  found_emails = []
  attachment_dir = 'attachments'
  
  # Create directory if it does'nt exist
  if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)
  
  print(f"Length of email id's {len(email_ids)}")
  
  for mail_id in email_ids:
    status, data = mail.fetch(mail_id, "(RFC822)") # Fetch email content
    raw_email = data[0][1] # The actual email
    
    # Parse the email
    msg = email.message_from_bytes(raw_email)
    subject, encoding = decode_header(msg["Subject"])[0]
    
    if isinstance(subject, bytes):
      subject = subject.decode(encoding or "utf-8")
      
    sender = msg.get("From")
    recipient = msg.get("To")
    
    # Get the date the email was recieved and parse it using paresedate_to_datetime
    received_date = msg.get("Date")
    received_date = email.utils.parsedate_to_datetime(received_date)
    
    # Convert the date to a standard date format (e.g. DD-MM-YYYY)
    received_date_str = received_date.strftime('%d-%m-%Y')
    
    # Check for attachments
    if msg.is_multipart():
      for part in msg.walk():
        if part.get_content_disposition() == "attachment":
          attachment = part.get_filename()
          if attachment and attachment.lower().endswith('.pdf'):
            # Check if this email has already been logged
            if is_email_logged(subject, attachment, received_date_str):
              print(f"Email with subject '{subject}' and attachment '{attachment}', and date '{received_date}' is already logged. Skipping...")
              continue
            
            # Save the attachment if it's a PDF
            attachment_data = part.get_payload(decode=True)
            file_path = os.path.join(attachment_dir, attachment)
            with open(file_path, 'wb') as f:
              f.write(attachment_data)
            print(f"Saved PDF: {attachment} to {file_path}")
            
            # Log the email
            email_data = {
              "sender": sender,
              "recipient": recipient,
              "subject": subject,
              "attachment": attachment,
              "received_date": received_date_str # Store in DD-MM-YYYY format
            }
            log_email(email_data)
            found_emails.append(email_data)
  
  mail.logout()
  return found_emails

# Function to send an email
def send_email(username, password, recipient, subject, body, attachment, attachment_path=None):
  # Set up the MIME
  msg = MIMEMultipart()
  msg["From"] = username
  msg["To"] = recipient
  msg["subject"] = subject
  
  # Add body to the email
  msg.attach(MIMEText(body, "plain"))
  
  # Attach file if provided
  if attachment_path:
    with open(attachment_path, "rb") as attachment:
      part = MIMEBase("application", "octet-stream")
      part.set_payload(attachment.read())
      encoders.encode_base64(part)
      part.add_header(
        "Content-Disposition",
        f"attachment; filename= {os.path.basename(attachment_path)}",
      )
      msg.attach(part)
  
  # Setup SMTP server
  server = smtplib.SMTP_SSL("YOUR SMTP DETAILS HERE", 465)
  server.login(username, password)
  
  # send the email
  server.sendmail(username, recipient, msg.as_string())

  # Quit the server
  server.quit()
  
def process_email(username, password, recipient_email):
  found_emails = search_emails(username, password, recipient_email)
  
  if found_emails:
    for email_data in found_emails:
      print(f"Email found: {email_data}")
      log_email(email_data)
      
  else:
    print("No email found, sending new one.")
    # Send a new email with or without an attachment
    # send_email(username, password, recipient_email, "Invoice", "Find email not previously sent.", "./attachments")

emails_with_attachments = search_emails("YOUR EMAIL ADDRESS HERE", "YOUR EMAIL PASSWORD", "SENDER EMAIL ADDRESS")


print(f"Found emails variable type is: {type(emails_with_attachments)}")

for email in emails_with_attachments:
    print(f"Found Email - Subject: {email['subject']}, Attachment: {email['attachment']}")

