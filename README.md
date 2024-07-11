
# Email Accessor and Processor

This Python script enables automated access to Gmail using IMAP, fetches specific emails based on criteria, processes attachments, and logs actions to a file.

## Features

- **IMAP Access**: Connects to Gmail's IMAP server securely using SSL.
- **Email Search**: Searches for unread emails within a specified date range and matching specific subject keywords.
- **Attachment Processing**: Downloads attachments from matching emails to a specified directory after sanitizing filenames.
- **Logging**: Logs processing actions and skip reasons to a log file for audit and debugging purposes.

## Requirements

- Python 3.6 or higher
- `imaplib`, `email`, `os`, `yaml`, `datetime`, `re`, `logging` (standard libraries)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/email-accessor.git
   cd email-accessor
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up configuration:
   - Rename `config.yml.example` to `config.yml`.
   - Edit `config.yml` with your Gmail account credentials (`user` and `password`).

## Configuration

Ensure the `config.yml` file is properly configured with your Gmail account credentials:

```yaml
email:
  user: your-email@gmail.com
  password: your-password
```

## Usage

Run the script `email_processor.py`:

```bash
python email_processor.py
```

### Script Workflow

1. **Login**: Connects to Gmail IMAP server securely.
2. **Fetch Emails**: Retrieves unread emails within the specified date range and matching subject keywords.
3. **Process Emails**:
   - Checks each email for attachments.
   - Downloads attachments to a specified directory (`attachments/`).
   - Logs processing actions (`ICI-Primary-Layout.log`).
   - Marks processed emails with a label `ProcessedByBOT` to avoid re-processing.

4. **Logout**: Closes the IMAP connection.

### Logging

All actions, including successful attachment downloads and skipped emails, are logged to `EMAILSCRAPPER.log` for review.

### Notes

- Ensure that your Gmail account allows access via less secure apps or generates an app-specific password if two-factor authentication is enabled.
- Modify `subject_keywords` in `email_processor.py` to match your specific email subject criteria.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

