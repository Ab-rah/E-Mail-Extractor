import imaplib
import email
from email.header import decode_header
import os
import yaml

# Function to decode email header
def decode_subject_header(header):
    decoded_header = decode_header(header)
    decoded_subject = []
    for part, encoding in decoded_header:
        if isinstance(part, bytes):
            decoded_subject.append(part.decode(encoding or 'utf-8'))
        else:
            decoded_subject.append(part)
    return ''.join(decoded_subject)


attachment_dir = "attachments"
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)


def read_credentials(yml_file):
    with open(yml_file,'r') as file:
        config = yaml.safe_load(file)
    return config['email']['user'],config['email']['password']

yml_file = 'config.yml'

email_address,password = read_credentials(yml_file)


imap_server = 'imap.gmail.com'
imap_port = 993
try:
    mail = imaplib.IMAP4_SSL(imap_server, imap_port)
    mail.login(email_address, password)
    mail.select('inbox')
    status, data = mail.search(None, 'ALL')


    for num in data[0].split():
        status, data = mail.fetch(num, '(RFC822)')
        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        for part in msg.walk():
            # Check if the part is an attachment
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()

            # Decode filename if it's encoded
            if filename:
                filename = decode_subject_header(filename)
                filepath = os.path.join(attachment_dir, filename)

                with open(filepath, 'wb') as f:
                    f.write(part.get_payload(decode=True))
                    print(f"Downloaded {filename} to {filepath}")
    # mail.logout()

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    try:
        mail.logout()
    except:
        pass
