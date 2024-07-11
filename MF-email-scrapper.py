import imaplib
import email
import os
import yaml
from datetime import datetime, timedelta
import re
import logging

logpath = r"C:\Users\AbdhulRahimSheikh.M\PycharmProjects\pythonProject\EmailAccessor\log\ICI-Primary-Layout.log"
logging.basicConfig(filename=logpath, level=logging.INFO, format='%(message)s')


def getCurrentTimestamp():
    current_date = datetime.today()
    next_date = current_date + timedelta(days=1)
    currentTimestamp = next_date.strftime('%d-%b-%Y')
    return currentTimestamp


def sanitize_filename(filename):
    # Replace invalid characters in the File Name with an underscore
    return re.sub(r'[\\/*?:"<>|\r\n]', '_', filename)


def read_credentials(yml_file):
    with open(yml_file, 'r') as file:
        config = yaml.safe_load(file)
    return config['email']['user'], config['email']['password']


def fetch_unprocessed_emails(imap_conn):
    start_date = '01-Dec-2023'
    end_date = getCurrentTimestamp()

    search_criteria_unread = 'UNSEEN'
    search_criteria_date = f'SINCE "{start_date}" BEFORE "{end_date}"'  # Ensure dates are properly formatted
    search_criteria = f'({search_criteria_unread} {search_criteria_date})'

    _, data = imap_conn.search(None, search_criteria)
    mail_id_list = data[0].split()

    msgs = []
    for num in mail_id_list:
        typ, data = imap_conn.fetch(num, '(BODY.PEEK[])')  # Use BODY.PEEK[] to avoid marking as read
        msgs.append((num, data))

    return msgs



def process_emails(imap_conn, msgs):
    subject_keywords = ["2023 ICI PRIMARY LAYOUT", "ICI PRIMARY LAYOUT 2023", "2023 PRIMARY Layout", "ICI 2023 Layout"]
    attachment_dir = r"C:\Users\AbdhulRahimSheikh.M\PycharmProjects\pythonProject\EmailAccessor\attachments"
    if not os.path.exists(attachment_dir):
        os.makedirs(attachment_dir)

    for num, msg in msgs[::-1]:
        for response_part in msg:
            if isinstance(response_part, tuple):
                attachmentFound = False
                my_msg = email.message_from_bytes(response_part[1])
                subject = my_msg['subject']
                subject_lower = subject.lower()
                if any(keyword.lower() in subject_lower for keyword in subject_keywords):
                    processed_labels = imap_conn.fetch(num, '(X-GM-LABELS)')
                    labels = processed_labels[1][0].decode('utf-8')
                    if 'ProcessedByBOT' in labels:
                        logging.info("|| SKIPPED: EMAIL ALREADY PROCESSED ||")
                        logging.info("Subject: {}".format(my_msg['subject']))
                        logging.info("From: {}".format(my_msg['from']))
                        logging.info(
                            "________________________________________________________________________________________________")
                        continue

                    for part in my_msg.walk():
                        if part.get_content_type() == 'text/plain':
                            pass
                        elif part.get_content_disposition() and part.get_content_disposition().startswith(
                                "attachment"):
                            filename = part.get_filename()
                            if filename:
                                attachmentFound = True
                                sanitized_filename = sanitize_filename(filename)
                                filepath = os.path.join(attachment_dir, sanitized_filename)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                logging.info(f"|| PROCESSED : ATTACHMENT HAS BEEN SAVED ||")
                                logging.info("Subject: {}".format(my_msg['subject']))
                                logging.info("From: {}".format(my_msg['from']))
                                logging.info(f"ATTACHMENT: {sanitized_filename}")
                                logging.info(
                                    "________________________________________________________________________________________________")
                    if attachmentFound:
                        label = 'ProcessedByBOT'
                        imap_conn.store(num, '+X-GM-LABELS', f'({label})')
                    else:
                        logging.info("|| SKIPPED: SUBJECT MATCHES BUT NO ATTACHMENT FOUND ||")
                        logging.info("Subject: {}".format(my_msg['subject']))
                        logging.info("From: {}".format(my_msg['from']))
                        logging.info(
                            "________________________________________________________________________________________________")
                else:
                    logging.info("|| SKIPPED: SUBJECT DOESN'T MATCH OUR CRITERIA ||")
                    logging.info("Subject: {}".format(my_msg['subject']))
                    logging.info("From: {}".format(my_msg['from']))
                    logging.info(
                        "________________________________________________________________________________________________")



def main():
    yml_file = 'config.yml'
    user, password = read_credentials(yml_file)
    imap_url = 'imap.gmail.com'
    my_mail = imaplib.IMAP4_SSL(imap_url)
    my_mail.login(user, password)
    my_mail.select('Inbox')

    # Fetch unread emails that do not have the 'ProcessedByBOT' label
    msgs = fetch_unprocessed_emails(my_mail)

    # Process the fetched emails
    process_emails(my_mail, msgs)

    print("Script Completed!")
    my_mail.logout()


if __name__ == "__main__":
    main()