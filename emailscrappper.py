import imaplib
import email
import os
import yaml

def read_credentials(yml_file):
    with open(yml_file,'r') as file:
        config = yaml.safe_load(file)
    return config['email']['user'],config['email']['password']

yml_file = 'config.yml'

user,password = read_credentials(yml_file)


imap_url = 'imap.gmail.com'
my_mail = imaplib.IMAP4_SSL(imap_url)
my_mail.login(user, password)

my_mail.select('Inbox')

key = 'FROM'
value = '123014001@sastra.ac.in'
_, data = my_mail.search(None, key, value)

mail_id_list = data[0].split()

msgs = []
for num in mail_id_list:
    typ, data = my_mail.fetch(num, '(RFC822)')  # RFC822 returns whole message (BODY fetches just body)
    msgs.append(data)

attachment_dir = "attachments"
if not os.path.exists(attachment_dir):
    os.makedirs(attachment_dir)

for msg in msgs[::-1]:
    for response_part in msg:
        if isinstance(response_part, tuple):
            my_msg = email.message_from_bytes(response_part[1])
            print("_________________________________________")
            print("Subject:", my_msg['subject'])
            print("From:", my_msg['from'])
            print("Body:")
            for part in my_msg.walk():
                if part.get_content_type() == 'text/plain':
                    print(part.get_payload())
                elif part.get_content_disposition() and part.get_content_disposition().startswith("attachment"):
                    filename = part.get_filename()
                    if filename:
                        filepath = os.path.join(attachment_dir, filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        print(f"Attachment saved: {filepath}")
# import os
# import yaml
# import imaplib
# import email
#
#
# def read_credentials(yml_file):
#     with open(yml_file, 'r') as file:
#         config = yaml.safe_load(file)
#     return config['email']['user'], config['email']['password']
#
#
# def login_to_email(imap_url, user, password):
#     my_mail = imaplib.IMAP4_SSL(imap_url)
#     my_mail.login(user, password)
#     return my_mail
#
#
# def search_emails(mail, key, value):
#     mail.select('Inbox')
#     _, data = mail.search(None, key, value)
#     return data[0].split()
#
#
# def fetch_emails(mail, mail_id_list):
#     msgs = []
#     for num in mail_id_list:
#         typ, data = mail.fetch(num, '(RFC822)')
#         msgs.append(data)
#     return msgs
#
#
# def save_attachments(msg, attachment_dir):
#     for response_part in msg:
#         if isinstance(response_part, tuple):
#             my_msg = email.message_from_bytes(response_part[1])
#             for part in my_msg.walk():
#                 if part.get_content_disposition() and part.get_content_disposition().startswith("attachment"):
#                     filename = part.get_filename()
#                     if filename:
#                         filepath = os.path.join(attachment_dir, filename)
#                         with open(filepath, "wb") as f:
#                             f.write(part.get_payload(decode=True))
#                         print(f"Attachment saved: {filepath}")
#
#
# def print_email_details(msg):
#     for response_part in msg:
#         if isinstance(response_part, tuple):
#             my_msg = email.message_from_bytes(response_part[1])
#             print("_________________________________________")
#             print("Subject:", my_msg['subject'])
#             print("From:", my_msg['from'])
#             print("Body:")
#             for part in my_msg.walk():
#                 if part.get_content_type() == 'text/plain':
#                     print(part.get_payload())
#
#
# def main(yml_file, imap_url, search_key, search_value, attachment_dir):
#     user, password = read_credentials(yml_file)
#     my_mail = login_to_email(imap_url, user, password)
#
#     if not os.path.exists(attachment_dir):
#         os.makedirs(attachment_dir)
#
#     mail_id_list = search_emails(my_mail, search_key, search_value)
#     msgs = fetch_emails(my_mail, mail_id_list)
#
#     for msg in msgs[::-1]:
#         print_email_details(msg)
#         save_attachments(msg, attachment_dir)
#
#
# if __name__ == "__main__":
#     yml_file = 'config.yml'
#     imap_url = 'imap.gmail.com'
#     search_key = 'FROM'
#     search_value = '123014001@sastra.ac.in'
#     attachment_dir = "attachments"
#
#     main(yml_file, imap_url, search_key, search_value, attachment_dir)
