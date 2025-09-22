import os
import logging
import pandas as pd
from datetime import datetime, timedelta
from pathlib import Path
from email.message import EmailMessage
import smtplib
import re

# --- Logging Setup ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')

# --- Environment Variables ---
SENDER_EMAIL = os.environ.get("SENDER_EMAIL", "").strip()
APP_PASSWORD = os.environ.get("APP_PASSWORD", "").strip()
RECIPIENTS = ["ganeshsai@nuevostech.com"]

if not SENDER_EMAIL:
    logging.warning("SENDER_EMAIL is not set or empty. Email functionality will be disabled.")
if not APP_PASSWORD:
    logging.warning("APP_PASSWORD is not set or empty. Email functionality will be disabled.")
logging.info(f"Using sender email: {SENDER_EMAIL}")
logging.info(f"Using recipients: {RECIPIENTS}")

# --- Template ---
TEMPLATE_SUBJECT = "Renewal Reminder"
TEMPLATE_BODY = """Hello {Client},

This is a reminder that your agreement is set to expire on {new_expiry}.

There are {days_left} day(s) left. Please take necessary action.

File path: {file_path}

Regards,
Renewals Team
"""

# --- Helpers ---
def is_valid_email(email):
    return bool(re.match(r"[^@]+@[^@]+\.[^@]+", email))

def build_message(sender, to_list, subject, body, attachment_path=None):
    if not sender or not is_valid_email(sender):
        logging.error("Sender email is empty or invalid.")
        return None

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_path and Path(attachment_path).exists():
        with open(attachment_path, "rb") as f:
            file_data = f.read()
            file_name = Path(attachment_path).name
            msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
    return msg

def send_email(msg):
    if not msg:
        return False
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        return True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return False

# --- Email Senders ---
def send_renewal_reminder(agreement, days_left):
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning(f"Email credentials missing. Skipping reminder for '{agreement['file']}'")
        return

    client_display = agreement.get('name') or agreement.get('email') or "Client"
    subject = f"Renewal Reminder: '{agreement['file']}' Expires Soon"
    body = TEMPLATE_BODY.format(
        Client=client_display,
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=days_left,
        file_path=agreement.get('path') or "N/A"
    )
    to_addr = [agreement.get('email')] if agreement.get('email') else RECIPIENTS
    attachment_path = agreement.get('path') if agreement.get('path') and Path(agreement['path']).exists() else None

    msg = build_message(SENDER_EMAIL, to_addr, subject, body, attachment_path)
    if msg and send_email(msg):
        logging.info(f"Reminder sent for '{agreement['file']}' ({days_left} days left) to {to_addr}")

def send_hourly_alert(agreement):
    if not SENDER_EMAIL or not APP_PASSWORD:
        logging.warning(f"Email credentials missing. Skipping hourly alert for '{agreement['file']}'")
        return

    client_display = agreement.get('name') or agreement.get('email') or "Client"
    subject = f"Renewal Reminder: '{agreement['file']}' Expires Today"
    body = TEMPLATE_BODY.format(
        Client=client_display,
        new_expiry=agreement['expiry_date'].strftime('%Y-%m-%d'),
        days_left=0,
        file_path=agreement.get('path') or "N/A"
    )
    to_addr = [agreement.get('email')] if agreement.get('email') else RECIPIENTS
    attachment_path = agreement.get('path') if agreement.get('path') and Path(agreement['path']).exists() else None

    msg = build_message(SENDER_EMAIL, to_addr, subject, body, attachment_path)
    if msg and send_email(msg):
        logging.info(f"Hourly alert sent for '{agreement['file']}' to {to_addr}")

# --- Agreement Logic ---
def parse_agreements_from_excel(file_path):
    df = pd.read_excel(file_path)
    df.columns = [c.strip().title() for c in df.columns]
    logging.info(f"Detected columns: {', '.join(df.columns)}")

    agreements = []
    for _, row in df.iterrows():
        try:
            agreement = {
                'name': row.get('Name', '').strip(),
                'email': row.get('Email', '').strip(),
                'expiry_date': pd.to_datetime(row['Expiry Date']).date(),
                'file': row.get('Name', '').strip(),
                'path': row.get('Path') if 'Path' in row else None
            }
            agreements.append(agreement)
        except Exception as e:
            logging.warning(f"Skipping row due to error: {e}")
    return agreements

def run_reminders_and_alerts(agreements):
    today = datetime.today().date()
    for agreement in agreements:
        days_left = (agreement['expiry_date'] - today).days
        if days_left < 0:
            continue
        elif days_left == 0:
            send_hourly_alert(agreement)
        elif days_left <= 5:
            send_renewal_reminder(agreement, days_left)

# --- Main Execution ---
def main_run_once():
    excel_path = "Renewal.xlsx"
    agreements = parse_agreements_from_excel(excel_path)
    logging.info(f"Parsed {len(agreements)} agreement(s).")
    for a in agreements[:10]:
        logging.info(f"  - {a['file']} | {a['expiry_date']} | {a['email']} | name={a['name']}")
    run_reminders_and_alerts(agreements)

if __name__ == "__main__":
    main_run_once()






