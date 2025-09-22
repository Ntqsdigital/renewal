

import os
import logging
import smtplib
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd

# Try to import gdown (better Google Drive handling); fallback to requests
try:
    import gdown
    _HAS_GDOWN = True
except Exception:
    import requests
    _HAS_GDOWN = False

# ---------- logging ----------
logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')

# ---------- env / config ----------
APP_EMAIL = (os.getenv('APP_EMAIL') or "").strip()
APP_PASSWORD = (os.getenv('APP_PASSWORD') or "").strip()
RECIPIENT_DEFAULT = (os.getenv('RECIPIENT_DEFAULT') or "").strip()
DOWNLOAD_PATH = Path(os.getenv('DOWNLOAD_PATH', 'Renewal.xlsx'))
SMTP_HOST = os.getenv('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))

if not APP_EMAIL:
    logging.warning("APP_EMAIL is not set. Emails will not be sent unless this is provided.")
if not APP_PASSWORD:
    logging.warning("APP_PASSWORD is not set or empty. Emails will not be sent unless this is provided.")

# If RECIPIENT_DEFAULT not set, we'll fall back to APP_EMAIL later when needed
RECIPIENT_DEFAULTS = [e.strip() for e in RECIPIENT_DEFAULT.split(',') if e.strip()] if RECIPIENT_DEFAULT else []
if not RECIPIENT_DEFAULTS and APP_EMAIL:
    RECIPIENT_DEFAULTS = [APP_EMAIL]

# ---------- helpers ----------
def download_with_gdown(drive_id: str, dest: Path) -> None:
    url = f"https://drive.google.com/uc?id={drive_id}&export=download"
    logging.info(f"Downloading via gdown: {url} -> {dest}")
    gdown.download(url, str(dest), quiet=False)

def download_with_requests(url: str, dest: Path) -> None:
    logging.info(f"Downloading via requests: {url} -> {dest}")
    resp = requests.get(url, stream=True)
    resp.raise_for_status()
    with open(dest, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)

def download_excel_from_gdrive(gdrive_id: str, dest: Path) -> None:
    """Download a Google Drive file ID to dest. Tries gdown first, then requests fallback."""
    dest_parent = dest.parent
    if dest_parent and not dest_parent.exists():
        dest_parent.mkdir(parents=True, exist_ok=True)

    if _HAS_GDOWN:
        try:
            download_with_gdown(gdrive_id, dest)
            if dest.exists() and dest.stat().st_size > 0:
                logging.info("Download complete (gdown).")
                return
            else:
                logging.warning("gdown reported success but file empty; falling back to requests.")
        except Exception as e:
            logging.warning(f"gdown download failed: {e} — will try requests fallback.")

    # requests fallback (use export=download url)
    try:
        url = f"https://drive.google.com/uc?id={gdrive_id}&export=download"
        # requests was imported earlier as fallback for gdown import failure
        if 'requests' not in globals():
            import requests
        download_with_requests(url, dest)
        logging.info("Download complete (requests fallback).")
    except Exception as e:
        logging.exception("Failed to download Excel from Google Drive.")
        raise

def find_columns(df: pd.DataFrame):
    """Return tuple (name_col, email_col, expiry_col). Uses case-insensitive fuzzy matching."""
    lc = [c.lower().strip() for c in df.columns]
    name_col = next((c for c in df.columns if any(tok in c.lower() for tok in ('name','client','person'))), None)
    email_col = next((c for c in df.columns if 'email' in c.lower()), None)
    expiry_col = next((c for c in df.columns if any(tok in c.lower() for tok in ('expiry','expiry date','due date','expires','end'))), None)

    # If not found by simple checks, try presence of date-like values in columns
    if not expiry_col:
        for c in df.columns:
            try:
                # try parse first non-null value
                sample = df[c].dropna().iloc[0] if len(df[c].dropna()) else None
                if sample is not None:
                    pd_ts = pd.to_datetime(sample, errors='coerce')
                    if not pd.isna(pd_ts):
                        expiry_col = c
                        logging.info(f"Selected '{c}' as expiry column based on content.")
                        break
            except Exception:
                continue

    return name_col, email_col, expiry_col

def send_email(subject: str, body: str, recipients: list) -> bool:
    """Send email using SMTP. Returns True on success, False on failure or if creds missing."""
    recipients = [r for r in recipients if r and isinstance(r, str)]
    if not recipients:
        logging.warning("No recipients provided; skipping email send.")
        return False
    if not (APP_EMAIL and APP_PASSWORD):
        logging.warning("Missing APP_EMAIL or APP_PASSWORD; skipping email send.")
        return False

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = APP_EMAIL
    msg['To'] = ", ".join(recipients)

    try:
        # Use context-aware starttls
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(APP_EMAIL, APP_PASSWORD)
            server.sendmail(APP_EMAIL, recipients, msg.as_string())
        logging.info(f"Email sent to: {recipients} | Subject: {subject}")
        return True
    except smtplib.SMTPAuthenticationError:
        logging.error("SMTP authentication failed. Check APP_EMAIL / APP_PASSWORD.")
        return False
    except Exception as e:
        logging.exception(f"Failed to send email: {e}")
        return False

# ---------- main ----------
def main():
    # Google Drive file id (from your URL)
    DRIVE_FILE_ID = "1aEyOe-C98I_sV0AItEewMFBl1l5R85R2"

    try:
        download_excel_from_gdrive(DRIVE_FILE_ID, DOWNLOAD_PATH)
    except Exception:
        logging.error("Could not download Excel file; exiting.")
        return

    # Load Excel. Let pandas infer header row (first row)
    try:
        df = pd.read_excel(DOWNLOAD_PATH, engine="openpyxl")
    except Exception as e:
        logging.exception(f"Failed to read Excel file: {e}")
        return

    logging.info("Excel Preview (first 8 rows):\n" + df.head(8).to_string(index=False))

    # Detect columns
    name_col, email_col, expiry_col = find_columns(df)
    logging.info(f"Detected columns -> Name: {name_col}, Email: {email_col}, Expiry: {expiry_col}")
    if not expiry_col:
        logging.error("Could not detect an expiry/date column. Ensure Excel has a date column.")
        return
    if not email_col:
        logging.warning("Could not detect an email column. Will use default recipients for notifications.")

    # Normalize expiry to datetime.date
    try:
        df[expiry_col] = pd.to_datetime(df[expiry_col], errors='coerce')
    except Exception:
        logging.exception("Error converting expiry column to datetime.")
        return

    # Filter rows where expiry is today
    today = pd.Timestamp.now().normalize().date()
    # create a date-only column for comparison
    df['_expiry_date_only'] = df[expiry_col].dt.date
    expiring_today = df[df['_expiry_date_only'] == today]

    logging.info(f"{len(expiring_today)} agreement(s) expiring today ({today}).")

    # recipients fallback: if a row email is missing, use RECIPIENT_DEFAULTS
    for idx, row in expiring_today.iterrows():
        person = row.get(name_col) if name_col else None
        person = str(person).strip() if person and not pd.isna(person) else "Client"
        to_email = None
        if email_col:
            raw_email = row.get(email_col)
            if pd.notna(raw_email) and str(raw_email).strip():
                to_email = str(raw_email).strip()

        recipients = [to_email] if to_email else RECIPIENT_DEFAULTS
        subject = f"Renewal Reminder for {person}"
        expiry_date_display = row[expiry_col].strftime("%Y-%m-%d") if pd.notna(row[expiry_col]) else "Unknown date"
        body = (
            f"Dear {person},\n\n"
            f"Your agreement expires today ({expiry_date_display}). Please take necessary action.\n\n"
            "Regards,\n"
            "NTQS Digital"
        )
        if not recipients:
            logging.warning(f"No recipients to notify for row {idx} ({person}); skipping.")
            continue

        send_email(subject, body, recipients)

if __name__ == "__main__":
    main()











