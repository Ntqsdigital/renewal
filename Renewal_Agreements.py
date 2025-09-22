import os
import gdown
import pandas as pd
from datetime import datetime, date
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
import logging
from zipfile import BadZipFile

# === CONFIGURATION ===
SENDER_EMAIL = os.getenv("APP_EMAIL", "ganeshsai@nuevostech.com")
APP_PASSWORD = os.getenv("APP_PASSWORD", "myrvnqpycpouccwb")
RECIPIENTS = [os.getenv("RECIPIENT_DEFAULT", "")]
FILE_ID = os.getenv("FILE_ID") or "1aEyOe-C98I_sV0AItEewMFBl1l5R85R2"
DOWNLOAD_PATH = Path(os.getenv("DOWNLOAD_PATH", "Renewal.xlsx"))
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))
SEND_CONFIRMATION = False
LOG_FILE = Path(os.getenv("LOG_FILE", "renewal.log"))

# === LOGGING SETUP ===
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

# === EMAIL TEMPLATE ===
TEMPLATE_SUBJECT = 'Renewed Expiry Date Notification'
TEMPLATE_BODY = """Hi {Client},

Your product's expiry date has been renewed.
New Expiry Date: {new_expiry}

• Attached File Path: {file_path}

Please take note and let us know if you have any questions.

Regards,
NTQS Digital

Note: Your agreement expires in {days_left} day(s)
"""

# === UTILITY FUNCTIONS ===
def download_from_drive(file_id: str, dest: Path):
    if not file_id.strip():
        logging.error("No FILE_ID provided. Set it via environment variable or config.")
        raise ValueError("FILE_ID is empty or not provided.")
    
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    logging.info(f"Downloading from: {url} -> {dest}")
    try:
        gdown.download(url, str(dest), quiet=False)
        logging.info("Download complete.")
    except Exception as e:
        logging.exception("Failed to download file from Google Drive.")
        raise

def is_valid_excel(path: Path) -> bool:
    try:
        with open(path, 'rb') as f:
            header = f.read(4)
        return header.startswith(b'PK')  # .xlsx is a zip file (starts with PK)
    except Exception:
        return False

def detect_header_and_load(path: Path) -> pd.DataFrame:
    try:
        raw = pd.read_excel(path, header=None, engine="openpyxl")
    except BadZipFile:
        raise ValueError("Downloaded file is not a valid .xlsx file.")
    except Exception:
        raise

    header_idx = None
    search_tokens = {"expiry", "expiry date", "email", "name", "file", "due", "end", "expires", "expires on", "due date", "client", "contact"}
    for i in range(min(50, len(raw))):
        row_vals = [str(x).strip().lower() for x in raw.iloc[i].fillna("")]
        combined = " ".join(row_vals)
        if any(tok in combined for tok in search_tokens):
            header_idx = i
            logging.info(f"Detected header at row {header_idx}: {row_vals}")
            break

    if header_idx is None:
        logging.warning("Could not confidently detect header row; defaulting to row 0.")
        header_idx = 0

    df = pd.read_excel(path, header=header_idx, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df

def build_message(sender: str, to_list: list, subject: str, body: str, attachment_path: str = None) -> EmailMessage:
    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = ", ".join(to_list)
    msg['Subject'] = subject
    msg.set_content(body, subtype='plain', charset='utf-8')
    if attachment_path:
        path = Path(attachment_path)
        if path.exists() and path.is_file():
            with open(path, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=path.name)
            logging.info(f"Attached file '{path.name}' to email.")
        else:
            logging.warning(f"Attachment file not found: {attachment_path}")
    return msg

def send_email(msg: EmailMessage):
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.ehlo()
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        logging.info(f"Email sent to: {msg['To']} | Subject: {msg['Subject']}")
    except smtplib.SMTPAuthenticationError:
        logging.exception("SMTP authentication error. Check credentials.")
        raise
    except Exception:
        logging.exception("Failed to send email.")
        raise

def make_agreements_list(df: pd.DataFrame) -> list:
    agreements = []
    expiry_cols = [c for c in df.columns if any(tok in c.lower() for tok in ('expiry','due','end','expires'))]
    email_cols = [c for c in df.columns if 'email' in c.lower()]
    file_cols = [c for c in df.columns if 'file' in c.lower() or c.lower() == 'name' or 'file name' in c.lower()]
    path_cols = [c for c in df.columns if 'path' in c.lower()]
    name_cols = [c for c in df.columns if any(tok in c.lower() for tok in ('name','client','contact','customer','person'))]

    if not expiry_cols:
        msg = f"No expiry-like column found. Columns: {list(df.columns)}"
        logging.error(msg)
        raise KeyError(msg)

    expiry_col = expiry_cols[0]
    email_col = email_cols[0] if email_cols else None
    file_col = file_cols[0] if file_cols else None
    path_col = path_cols[0] if path_cols else None
    name_col = name_cols[0] if name_cols else None

    logging.info(f"Using columns: expiry='{expiry_col}', email='{email_col}', file='{file_col}', path='{path_col}', name='{name_col}'")

    for idx, row in df.iterrows():
        try:
            expiry_dt = pd.to_datetime(row.get(expiry_col), errors='coerce', dayfirst=True)
            if pd.isna(expiry_dt):
                continue

            display_name = str(row.get(file_col, '')).strip() or str(row.get(name_col, '')).strip() or str(row.get(email_col, '')).strip() or 'Unnamed Agreement'
            email_val = str(row.get(email_col, '')).strip() if email_col else ''
            name_val = str(row.get(name_col, '')).strip() if name_col else ''

            agreements.append({
                'file': display_name,
                'expiry_date': expiry_dt,
                'path': str(row.get(path_col, '')).strip() if path_col else '',
                'status': 'EXPIRES TODAY' if expiry_dt.date() == datetime.now().date() else 'Upcoming',
                'email': email_val if email_val else RECIPIENTS[0],
                'name': name_val
            })
        except Exception:
            logging.exception(f"Error parsing row {idx}; skipping.")
            continue
    return agreements

def send_confirmation_email(client_name: str = "Client", to_emails: list = None):
    to_addresses = to_emails if to_emails else RECIPIENTS
    body = TEMPLATE_BODY.format(Client=client_name, new_expiry="15-09-2025", days_left=0, file_path="N/A")
    msg = build_message(SENDER_EMAIL, to_addresses, TEMPLATE_SUBJECT, body)
    send_email(msg)

def send_renewal_reminder(agreement: dict, days_left: int):
    client_display = agreement.get('name') or agreement.get('email') or "Client"
    body = TEMPLATE_BODY.format(
        Client=client_display,
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=days_left,
        file_path=agreement['path'] or "N/A"
    )
    to_addr = [agreement.get('email')] if agreement.get('email') else RECIPIENTS
    msg = build_message(SENDER_EMAIL, to_addr, f"Renewal Reminder: '{agreement['file']}' Expires Soon", body, attachment_path=agreement['path'])
    send_email(msg)

def send_hourly_alert(agreement: dict):
    client_display = agreement.get('name') or agreement.get('email') or "Client"
    body = TEMPLATE_BODY.format(
        Client=client_display,
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=0,
        file_path=agreement['path'] or "N/A"
    )
    to_addr = [agreement.get('email')] if agreement.get('email') else RECIPIENTS
    msg = build_message(SENDER_EMAIL, to_addr, f"Renewal Reminder: '{agreement['file']}' Expires Today", body, attachment_path=agreement['path'])
    send_email(msg)

def run_reminders_and_alerts(agreements: list):
    today = date.today()
    for agreement in agreements:
        days_left = (agreement['expiry_date'].date() - today).days
        if 1 <= days_left <= 5:
            send_renewal_reminder(agreement, days_left)

    expiring_today = [a for a in agreements if a['expiry_date'].date() == today]
    if expiring_today:
        logging.info(f"{len(expiring_today)} agreement(s) expire today.")
        for agreement in expiring_today:
            send_hourly_alert(agreement)
    else:
        logging.info("No agreements expiring today.")

def main_run_once():
    try:
        download_from_drive(FILE_ID, DOWNLOAD_PATH)
    except Exception:
        logging.error("Download failed. Exiting.")
        return

    if not is_valid_excel(DOWNLOAD_PATH):
        logging.error("Downloaded file is not a valid Excel file. Aborting.")
        return

    try:
        df = detect_header_and_load(DOWNLOAD_PATH)
    except Exception as e:
        logging.exception("Failed to read Excel file.")
        return

    logging.info("Excel Preview:\n" + df.head(8).to_string(index=False))
    logging.info("Detected columns: " + ", ".join(list(df.columns)))

    try:
        agreements = make_agreements_list(df)
        logging.info(f"Parsed {len(agreements)} agreements.")
    except Exception:
        logging.exception("Failed to build agreements list.")
        return

    if SEND_CONFIRMATION and agreements:
        first = agreements[0]
        send_confirmation_email(client_name=first.get('name') or first.get('email'), to_emails=[first.get('email')])

    run_reminders_and_alerts(agreements)

if __name__ == "__main__":
    main_run_once()









