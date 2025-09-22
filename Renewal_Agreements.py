import os
import gdown
import pandas as pd
from datetime import datetime
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
import logging

# ------------- CONFIG -------------
# Environment variables
SENDER_EMAIL = os.getenv('APP_EMAIL', '').strip()
APP_PASSWORD = os.getenv('APP_PASSWORD', '').strip()
RECIPIENT_DEFAULT = os.getenv('RECIPIENT_DEFAULT', 'ganeshsai@nuevostech.com').strip()
FILE_ID = os.getenv('FILE_ID', '1aEyOe-C98I_sV0AItEewMFBl1l5R85R2').strip()
DOWNLOAD_PATH = Path(os.getenv('DOWNLOAD_PATH', 'Renewal.xlsx'))
SMTP_HOST = os.getenv('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', '587'))
LOG_FILE = Path(os.getenv('LOG_FILE', 'renewal.log'))

# Setup logging to file and console
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# ------------- FUNCTIONS -------------

def download_file_from_drive(file_id: str, dest_path: Path):
    if not file_id:
        logging.error("Google Drive FILE_ID is not set.")
        raise ValueError("Google Drive FILE_ID is required.")
    url = f'https://drive.google.com/uc?id={file_id}&export=download'
    logging.info(f"Downloading {url} to {dest_path} ...")
    try:
        gdown.download(url, str(dest_path), quiet=False)
        if not dest_path.exists() or dest_path.stat().st_size == 0:
            raise FileNotFoundError(f"Download failed or file empty: {dest_path}")
        logging.info("Download completed successfully.")
    except Exception as e:
        logging.exception("Failed to download the file.")
        raise e

def detect_header_row(df_raw: pd.DataFrame) -> int:
    """Try to detect header row index by scanning first 50 rows for common keywords"""
    keywords = ['expiry', 'email', 'name', 'file', 'due', 'end', 'expires', 'contact', 'client']
    max_rows = min(50, len(df_raw))
    for i in range(max_rows):
        row_str = " ".join(str(x).lower() for x in df_raw.iloc[i].fillna(''))
        if any(kw in row_str for kw in keywords):
            logging.info(f"Detected header row at index {i} with values: {df_raw.iloc[i].tolist()}")
            return i
    logging.warning("Header row could not be confidently detected. Using first row as header.")
    return 0

def load_excel_with_header(filepath: Path) -> pd.DataFrame:
    raw_df = pd.read_excel(filepath, header=None, engine='openpyxl')
    header_idx = detect_header_row(raw_df)
    df = pd.read_excel(filepath, header=header_idx, engine='openpyxl')
    df.columns = [str(c).strip() for c in df.columns]
    df.dropna(how='all', inplace=True)
    df.reset_index(drop=True, inplace=True)
    logging.info(f"Excel loaded with header row {header_idx}, columns: {list(df.columns)}")
    return df

def find_column(columns, keywords):
    for kw in keywords:
        for col in columns:
            if kw.lower() in col.lower():
                return col
    return None

def parse_agreements(df: pd.DataFrame):
    # Find relevant columns
    expiry_col = find_column(df.columns, ['expiry', 'due', 'end', 'expires'])
    email_col = find_column(df.columns, ['email', 'contact'])
    name_col = find_column(df.columns, ['name', 'client', 'person'])
    file_col = find_column(df.columns, ['file', 'filename'])
    path_col = find_column(df.columns, ['path'])

    if not expiry_col:
        logging.error("No expiry date column found. Exiting.")
        raise KeyError("Expiry date column required.")

    agreements = []
    for idx, row in df.iterrows():
        try:
            expiry_raw = row[expiry_col]
            expiry_date = pd.to_datetime(expiry_raw, errors='coerce')
            if pd.isna(expiry_date):
                logging.debug(f"Skipping row {idx} due to invalid expiry date: {expiry_raw}")
                continue

            agreement_name = str(row[file_col] if file_col and pd.notna(row[file_col]) else
                                 row[name_col] if name_col and pd.notna(row[name_col]) else
                                 f"Agreement #{idx+1}").strip()

            email = str(row[email_col]).strip() if email_col and pd.notna(row[email_col]) else RECIPIENT_DEFAULT
            email = email.replace('\n', '').replace('\r', '')

            path = str(row[path_col]).strip() if path_col and pd.notna(row[path_col]) else ''

            agreements.append({
                'name': agreement_name,
                'expiry_date': expiry_date,
                'email': email,
                'path': path
            })
        except Exception as e:
            logging.warning(f"Skipping row {idx} due to error: {e}")

    logging.info(f"Parsed {len(agreements)} agreements successfully.")
    return agreements

def build_email_message(sender, recipients, subject, body, attachment_path=None):
    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject
    msg.set_content(body)

    if attachment_path and Path(attachment_path).exists():
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = Path(attachment_path).name
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
        logging.info(f"Attached file {file_name} to email.")
    elif attachment_path:
        logging.warning(f"Attachment path provided but file not found: {attachment_path}")

    return msg

def send_email(msg: EmailMessage):
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning("Email credentials not set. Skipping email sending.")
        return False

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls(context=context)
            server.login(SENDER_EMAIL, APP_PASSWORD)
            server.send_message(msg)
        logging.info(f"Email sent to {msg['To']}")
        return True
    except smtplib.SMTPAuthenticationError:
        logging.error("SMTP Authentication failed. Check your email and app password.")
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
    return False

# ------------- MAIN LOGIC -------------

def main():
    logging.info("Script started.")

    # Download the latest renewal Excel
    try:
        download_file_from_drive(FILE_ID, DOWNLOAD_PATH)
    except Exception as e:
        logging.error(f"Cannot proceed without Excel file: {e}")
        return

    # Load and parse Excel
    try:
        df = load_excel_with_header(DOWNLOAD_PATH)
        agreements = parse_agreements(df)
    except Exception as e:
        logging.error(f"Error parsing Excel: {e}")
        return

    # For each agreement expiring today or soon, send email
    today = datetime.now().date()
    for ag in agreements:
        days_left = (ag['expiry_date'].date() - today).days
        if days_left < 0:
            continue  # already expired, ignore or handle differently
        if days_left > 7:
            continue  # ignore far future expiry for now

        # Compose email
        subject = f"Expiry Reminder for {ag['name']}"
        body = f"""Hi {ag['name']},

Your agreement expires in {days_left} day(s) on {ag['expiry_date'].date()}.

Please find the details attached.

Regards,
Your Team
"""
        msg = build_email_message(
            sender=SENDER_EMAIL or RECIPIENT_DEFAULT,
            recipients=[ag['email']],
            subject=subject,
            body=body,
            attachment_path=ag['path'] if ag['path'] else None
        )

        # Send email if credentials are set
        if send_email(msg):
            logging.info(f"Reminder sent for agreement '{ag['name']}' to {ag['email']}")
        else:
            logging.info(f"Skipped sending reminder for '{ag['name']}' due to missing credentials.")

    logging.info("Script execution finished.")

if __name__ == "__main__":
    main()






