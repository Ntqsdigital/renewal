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

# === Safe Credential Input ===
def _get_env_fallback(key: str, fallback: str = "") -> str:
    raw = os.getenv(key)
    if raw is None:
        return fallback
    trimmed = str(raw).strip()
    return trimmed if trimmed else fallback

# Optional local testing (interactive input)
if sys.stdin.isatty():
    try:
        import getpass
        user = input("Enter User ID: ")
        pwd = getpass.getpass("Enter Password: ")
        os.environ["USER_ID"] = user
        os.environ["APP_PASSWORD"] = pwd
        print(f"User: {user}, Password securely entered")
    except Exception as e:
        print(f"Input error: {e}")

# === Environment Config ===
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

# === Logging ===
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"), logging.StreamHandler()]
)

# === Email Templates ===
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

# === Utility ===
def clean_email_address(email: str) -> str:
    return email.strip().replace('\n', '').replace('\r', '') if email else ""

# === Download ===
def download_from_drive(file_id: str, dest: Path):
    if not file_id.strip():
        raise ValueError("FILE_ID is empty or missing.")
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    logging.info(f"Downloading {url} -> {dest} ...")
    try:
        gdown.download(url, str(dest), quiet=False)
        if not dest.exists() or dest.stat().st_size == 0:
            raise FileNotFoundError(f"Download failed or file is empty: {dest}")
        logging.info("Download complete.")
    except Exception:
        logging.exception("Failed to download file from Google Drive.")
        raise

# === Excel Header Detection ===
def detect_header_and_load(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")
    header_idx = None
    search_tokens = {"expiry", "expiry date", "email", "name", "file", "due", "end", "expires", "expires on", "due date", "client", "contact"}
    max_scan = min(50, len(raw))
    for i in range(max_scan):
        row_vals = [str(x).strip().lower() for x in raw.iloc[i].fillna("")]
        combined = " ".join(row_vals)
        if any(tok in combined for tok in search_tokens):
            header_idx = i
            logging.info(f"Detected header at row {header_idx} (0-based). Row values: {row_vals}")
            break
    if header_idx is None:
        header_idx = 0
        logging.warning("Could not confidently detect header row; using row 0 as header.")
    df = pd.read_excel(path, header=header_idx, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    return df

# === Email Message Builder ===
def build_message(sender: str, to_list: list, subject: str, body: str, attachment_path: str = None) -> EmailMessage:
    def _clean(e): return str(e).strip().replace('\n', '').replace('\r', '') if e else ""
    clean_sender = _clean(sender or RECIPIENTS[0])
    clean_to_list = [_clean(email) for email in (to_list or []) if _clean(email)]
    clean_subject = subject.strip().replace('\n', ' ').replace('\r', ' ') if subject else "No subject"

    msg = EmailMessage()
    msg['From'] = clean_sender
    msg['To'] = ", ".join(clean_to_list)
    msg['Subject'] = clean_subject
    msg.set_content(body or "", subtype='plain', charset='utf-8')

    if attachment_path:
        path = Path(attachment_path)
        if path.exists() and path.is_file():
            with open(path, 'rb') as f:
                msg.add_attachment(f.read(), maintype='application', subtype='octet-stream', filename=path.name)
            logging.info(f"Attached file '{path.name}' to email.")
        else:
            logging.warning(f"Attachment file not found: {attachment_path}")
    return msg

# === Email Sending ===
def send_email(msg: EmailMessage):
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning("Missing email credentials. Skipping email send.")
        return False
    context = ssl.create_default_context()
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        logging.info(f"Email sent to: {msg['To']} | Subject: {msg['Subject']}")
        return True
    except smtplib.SMTPAuthenticationError:
        logging.error("SMTP authentication error. Check app password or account security settings.")
    except Exception:
        logging.exception("Failed to send email.")
    return False

# === Agreement Parser ===
def make_agreements_list(df: pd.DataFrame) -> list:
    agreements = []
    expiry_cols = [c for c in df.columns if any(tok in c.lower() for tok in ('expiry','due','end','expires'))]
    email_cols = [c for c in df.columns if 'email' in c.lower()]
    file_cols = [c for c in df.columns if 'file' in c.lower() or c.lower() == 'name' or 'file name' in c.lower()]
    path_cols = [c for c in df.columns if 'path' in c.lower()]
    name_cols = [c for c in df.columns if any(tok in c.lower() for tok in ('name','client','contact','customer','person'))]

    if not expiry_cols:
        raise KeyError("No expiry-like column found in the Excel file.")

    expiry_col = expiry_cols[0]
    email_col = email_cols[0] if email_cols else None
    file_col = file_cols[0] if file_cols else None
    path_col = path_cols[0] if path_cols else None
    name_col = name_cols[0] if name_cols else None

    logging.info(f"Using expiry column: '{expiry_col}', email: '{email_col}', file: '{file_col}', path: '{path_col}', name: '{name_col}'")

    for idx, row in df.iterrows():
        try:
            expiry_raw = row.get(expiry_col)
            expiry_dt = pd.to_datetime(expiry_raw, errors='coerce', dayfirst=True)
            if pd.isna(expiry_dt):
                continue
            display_name = str(row.get(file_col, '') or row.get(name_col, '') or row.get(email_col, '') or 'Unnamed Agreement').strip()
            email_val = clean_email_address(str(row.get(email_col, ''))) if email_col else ''
            name_val = str(row.get(name_col, '')).strip() if name_col else ''
            item = {
                'file': display_name,
                'expiry_date': expiry_dt,
                'path': str(row.get(path_col, '')).strip() if path_col else '',
                'status': 'EXPIRES TODAY' if expiry_dt.date() == datetime.now().date() else 'Upcoming',
                'email': email_val or clean_email_address(RECIPIENTS[0]),
                'name': name_val
            }
            agreements.append(item)
        except Exception:
            logging.exception(f"Error parsing row {idx}; skipping.")
            continue
    return agreements

# === Email Logic ===
def send_confirmation_email(client_name: str = "Client", to_emails: list = None):
    to_addresses = to_emails if to_emails else RECIPIENTS
    subject = TEMPLATE_SUBJECT
    body = TEMPLATE_BODY.format(Client=client_name, new_expiry="15-09-2025", days_left=0, file_path="N/A")
    msg = build_message(SENDER_EMAIL, to_addresses, subject, body)
    send_email(msg)

def send_renewal_reminder(agreement: dict, days_left: int):
    client_display = agreement.get('name') or agreement.get('email') or "Client"
    subject = f"Renewal Reminder: '{agreement['file']}' Expires Soon"
    body = TEMPLATE_BODY.format(
        Client=client_display,
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=days_left,
        file_path=agreement['path'] or "N/A"
    )
    to_addr = [agreement.get('email')] or RECIPIENTS
    msg = build_message(SENDER_EMAIL, to_addr, subject, body, attachment_path=agreement.get('path'))
    send_email(msg)

def send_hourly_alert(agreement: dict):
    client_display = agreement.get('name') or agreement.get('email') or "Client"
    subject = f"Renewal Reminder: '{agreement['file']}' Expires Today"
    body = TEMPLATE_BODY.format(
        Client=client_display,
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=0,
        file_path=agreement['path'] or "N/A"
    )
    to_addr = [agreement.get('email')] or RECIPIENTS
    msg = build_message(SENDER_EMAIL, to_addr, subject, body, attachment_path=agreement.get('path'))
    send_email(msg)

# === Run Logic ===
def run_reminders_and_alerts(agreements: list):
    today = date.today()
    for agreement in agreements:
        days_left = (agreement['expiry_date'].date() - today).days
        if 1 <= days_left <= 5:
            send_renewal_reminder(agreement, days_left)
    todays = [a for a in agreements if a['expiry_date'].date() == today]
    for agreement in todays:
        send_hourly_alert(agreement)

# === Main ===
def main_run_once():
    logging.info(f"Using sender email: {SENDER_EMAIL}")
    logging.info(f"Using recipients: {RECIPIENTS}")
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning("Email sending is disabled due to missing credentials.")
    try:
        download_from_drive(FILE_ID, DOWNLOAD_PATH)
        df = detect_header_and_load(DOWNLOAD_PATH)
        logging.info("Excel loaded:\n" + df.head(8).to_string(index=False))
        agreements = make_agreements_list(df)
        logging.info(f"Parsed {len(agreements)} agreement(s).")
        if SEND_CONFIRMATION and agreements:
            send_confirmation_email(client_name=agreements[0].get('name', 'Client'), to_emails=[agreements[0].get('email')])
        run_reminders_and_alerts(agreements)
    except Exception:
        logging.exception("Fatal error in main_run_once")

    logging.info("Script execution completed.")

if __name__ == "__main__":
    main_run_once()










