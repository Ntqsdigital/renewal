import os
import logging
import requests
import smtplib
import pandas as pd
from email.message import EmailMessage

# ================== CONFIG ==================
FILE_ID = os.getenv("FILE_ID", "").strip()   # Google Drive File ID
DOWNLOAD_PATH = "agreements.xlsx"

SENDER_EMAIL = os.getenv("SENDER_EMAIL", "").strip()
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD", "").strip()
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
# ============================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)

# ---------- Download from Google Drive ----------
def download_from_drive(file_id, download_path):
    if not file_id:
        raise ValueError("FILE_ID is empty. Please set FILE_ID properly.")

    url = f"https://drive.google.com/uc?id={file_id}"
    logging.info(f"Downloading file from {url}")
    r = requests.get(url, allow_redirects=True)

    if r.status_code != 200:
        raise Exception(f"Failed to download file, status code: {r.status_code}")

    with open(download_path, "wb") as f:
        f.write(r.content)

    logging.info(f"Downloaded file saved to {download_path}")


# ---------- Build Email Safely ----------
def build_message(sender, to_addr, subject, body, attachment_path=None):
    msg = EmailMessage()

    # Strip hidden \n \r from headers
    sender = sender.strip()
    to_addr = to_addr.strip()
    subject = subject.strip()

    msg["From"] = sender
    msg["To"] = to_addr
    msg["Subject"] = subject

    msg.set_content(body)

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype="application",
                           subtype="octet-stream", filename=file_name)

    return msg


# ---------- Send Reminder ----------
def send_renewal_reminder(agreement, days_left):
    to_addr = str(agreement["email"])
    name = str(agreement["name"])
    subject = f"Renewal Reminder: {name} - {days_left} days left"
    body = f"""
Hello {name},

This is a reminder that your agreement will expire in {days_left} days.

Thank you.
"""

    msg = build_message(SENDER_EMAIL, to_addr, subject, body,
                        attachment_path=agreement.get("path"))

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)

    logging.info(f"Sent reminder to {to_addr}")


# ---------- Main Run ----------
def main_run_once():
    try:
        download_from_drive(FILE_ID, DOWNLOAD_PATH)

        # Load agreements from Excel
        df = pd.read_excel(DOWNLOAD_PATH)

        # Expected columns: name | email | days_left | path(optional)
        agreements = df.to_dict(orient="records")

        for ag in agreements:
            days_left = int(ag.get("days_left", 7))  # default 7 if missing
            send_renewal_reminder(ag, days_left)

    except Exception as e:
        logging.error(f"Process failed. Exiting. Error: {e}")
        raise


if __name__ == "__main__":
    main_run_once()






