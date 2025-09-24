import os
import sys
import gdown
import pandas as pd
from datetime import datetime, date
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
import logging
def _get_env_fallback(key: str, fallback: str = "") -> str:
    raw = os.getenv(key)
    return str(raw).strip() if raw else fallback
if sys.stdin.isatty():
    try:
        import getpass
        user = input("Enter User ID: ")
        pwd = getpass.getpass("Enter App Password: ")
        os.environ["APP_EMAIL"] = user
        os.environ["APP_PASSWORD"] = pwd
        print("Credentials set via terminal input.")
    except Exception as e:
        print(f"Input error: {e}")
SENDER_EMAIL = _get_env_fallback("APP_EMAIL", "ganeshsai@nuevostech.com")
APP_PASSWORD = _get_env_fallback("APP_PASSWORD", "")
recipient_default = _get_env_fallback("RECIPIENT_DEFAULT", "")
RECIPIENTS = [recipient_default] if recipient_default else ["ganeshsai@nuevostech.com"]
FILE_ID = _get_env_fallback("FILE_ID", "1aEyOe-C98I_sV0AItEewMFBl1l5R85R2")
DOWNLOAD_PATH = Path(_get_env_fallback("DOWNLOAD_PATH", "Renewal.xlsx"))
SMTP_HOST = _get_env_fallback("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(_get_env_fallback("SMTP_PORT", "587"))
SEND_CONFIRMATION = False
LOG_FILE = Path(_get_env_fallback("LOG_FILE", "renewal.log"))
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)
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
def clean_email_address(email: str) -> str:
    return email.strip().replace('\n', '').replace('\r', '') if email else ""
def download_from_drive(file_id: str, dest: Path):
    if not file_id.strip():
        raise ValueError("FILE_ID is empty or missing.")
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    logging.info(f"Downloading from: {url} -> {dest}")
    try:
        gdown.download(url, str(dest), quiet=False)
        if not dest.exists() or dest.stat().st_size == 0:
            raise FileNotFoundError("File downloaded but is empty.")
        logging.info(" Download complete.")
    except Exception as e:
        logging.error(" Failed to download file from Google Drive.")
        logging.error(" Ensure file is shared with 'Anyone with the link'.")
        logging.error(f"Google Drive URL: {url}")
        raise
def detect_header_and_load(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    header_idx = None
    search_tokens = {"expiry", "expiry date", "email", "name", "file", "due", "end", "expires", "expires on", "due date", "client", "contact"}
    for i in range(min(50, len(raw))):
        row_vals = [str(x).strip().lower() for x in raw.iloc[i].fillna("")]
        if any(tok in " ".join(row_vals) for tok in search_tokens):
            header_idx = i
            logging.info(f"Detected header at row {header_idx}: {row_vals}")
            break
    if header_idx is None:
        header_idx = 0
        logging.warning("No clear header found; defaulting to row 0.")
    df = pd.read_excel(path, header=header_idx, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df.dropna(how="all").reset_index(drop=True)
def build_message(sender: str, to_list: list, subject: str, body: str, attachment_path: str = None) -> EmailMessage:
    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = ", ".join(to_list)
    msg['Subject'] = subject
    msg.set_content(body, subtype='plain', charset='utf-8')
    if attachment_path:
        path = Path(attachment_path)
        if path.exists():
            with open(path, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=path.name)
            logging.info(f"Attached file: {path.name}")
        else:
            logging.warning(f"Attachment not found: {attachment_path}")
    return msg
def send_email(msg: EmailMessage):
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning("Email credentials missing. Skipping send.")
        return False
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        logging.info(f"📧 Email sent to: {msg['To']}")
        return True
    except smtplib.SMTPAuthenticationError:
        logging.error("SMTP Authentication failed. Check app password.")
    except Exception:
        logging.exception("Failed to send email.")
    return False
def make_agreements_list(df: pd.DataFrame) -> list:
    agreements = []
    expiry_cols = [c for c in df.columns if any(tok in c.lower() for tok in ('expiry','due','end','expires'))]
    email_cols = [c for c in df.columns if 'email' in c.lower()]
    file_cols = [c for c in df.columns if 'file' in c.lower() or c.lower() == 'name']
    path_cols = [c for c in df.columns if 'path' in c.lower()]
    name_cols = [c for c in df.columns if any(tok in c.lower() for tok in ('name','client','contact','customer'))]
    if not expiry_cols:
        raise KeyError("No expiry column found in the file.")
    expiry_col = expiry_cols[0]
    email_col = email_cols[0] if email_cols else None
    file_col = file_cols[0] if file_cols else None
    path_col = path_cols[0] if path_cols else None
    name_col = name_cols[0] if name_cols else None
    logging.info(f"Using columns - Expiry: {expiry_col}, Email: {email_col}, Name: {name_col}, File: {file_col}")
    for _, row in df.iterrows():
        expiry_dt = pd.to_datetime(row.get(expiry_col), errors='coerce', dayfirst=True)
        if pd.isna(expiry_dt):
            continue
        agreements.append({
            'file': str(row.get(file_col, '') or 'Unnamed').strip(),
            'expiry_date': expiry_dt,
            'email': clean_email_address(str(row.get(email_col, ''))),
            'name': str(row.get(name_col, '')).strip(),
            'path': str(row.get(path_col, '')).strip()
        })
    return agreements
def send_renewal_reminder(agreement: dict, days_left: int):
    subject = f"Renewal Reminder: {agreement['file']} Expires Soon"
    body = TEMPLATE_BODY.format(
        Client=agreement['name'] or 'Client',
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=days_left,
        file_path=agreement['path'] or 'N/A'
    )
    to = [agreement['email']] if agreement['email'] else RECIPIENTS
    msg = build_message(SENDER_EMAIL, to, subject, body, agreement['path'])
    send_email(msg)
def send_hourly_alert(agreement: dict):
    subject = f"Urgent: {agreement['file']} Expires Today"
    body = TEMPLATE_BODY.format(
        Client=agreement['name'] or 'Client',
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=0,
        file_path=agreement['path'] or 'N/A'
    )
    to = [agreement['email']] if agreement['email'] else RECIPIENTS
    msg = build_message(SENDER_EMAIL, to, subject, body, agreement['path'])
    send_email(msg)
def run_reminders_and_alerts(agreements: list):
    today = date.today()
    for ag in agreements:
        days_left = (ag['expiry_date'].date() - today).days
        if days_left == 0:
            send_hourly_alert(ag)
        elif 1 <= days_left <= 5:
            send_renewal_reminder(ag, days_left)
def main_run_once():
    logging.info(f"Using sender email: {SENDER_EMAIL}")
    logging.info(f"Using recipients: {RECIPIENTS}")
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning("Email sending is disabled due to missing credentials.")
    try:
        download_from_drive(FILE_ID, DOWNLOAD_PATH)
        df = detect_header_and_load(DOWNLOAD_PATH)
        agreements = make_agreements_list(df)
        logging.info(f" Parsed {len(agreements)} agreement(s).")
        run_reminders_and_alerts(agreements)
    except Exception:
        logging.exception("Fatal error during execution.")
    logging.info("Script completed")
if __name__ == "__main__":
    main_run_once()











