import os
import logging
import smtplib
from email.mime.text import MIMEText
import pandas as pd
import requests

logging.basicConfig(level=logging.INFO, format='%(asctime)s | %(levelname)s | %(message)s')

# Load environment variables
APP_EMAIL = os.getenv('APP_EMAIL')
APP_PASSWORD = os.getenv('APP_PASSWORD')
RECIPIENT_DEFAULT = os.getenv('RECIPIENT_DEFAULT')
DOWNLOAD_PATH = os.getenv('DOWNLOAD_PATH', 'Renewal.xlsx')
SMTP_HOST = os.getenv('SMTP_HOST', 'smtp.gmail.com')
SMTP_PORT = int(os.getenv('SMTP_PORT', 587))

if not APP_EMAIL:
    logging.error("APP_EMAIL is not set. Email functionality will be disabled.")
if not APP_PASSWORD:
    logging.warning("APP_PASSWORD is not set or empty. Email functionality will be disabled.")
if not RECIPIENT_DEFAULT:
    logging.info("RECIPIENT_DEFAULT not set. Using sender email as recipient if available.")
    RECIPIENT_DEFAULT = APP_EMAIL if APP_EMAIL else ''

RECIPIENTS = [email.strip() for email in RECIPIENT_DEFAULT.split(',')] if RECIPIENT_DEFAULT else []

def download_excel(url, path):
    logging.info(f"Downloading {url} -> {path} ...")
    response = requests.get(url)
    response.raise_for_status()
    with open(path, 'wb') as f:
        f.write(response.content)
    logging.info("Download complete.")

def send_email(subject, body, recipients):
    if not (APP_EMAIL and APP_PASSWORD and recipients):
        logging.warning("Email not sent due to missing email credentials or recipients.")
        return False
    try:
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = APP_EMAIL
        msg['To'] = ', '.join(recipients)

        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.starttls()
        server.login(APP_EMAIL, APP_PASSWORD)
        server.sendmail(APP_EMAIL, recipients, msg.as_string())
        server.quit()

        logging.info(f"Email sent to: {recipients}")
        return True
    except Exception as e:
        logging.error(f"Failed to send email: {e}")
        return False

def main():
    excel_url = 'https://drive.google.com/uc?id=1aEyOe-C98I_sV0AItEewMFBl1l5R85R2&export=download'

    # Download Renewal Excel file
    download_excel(excel_url, DOWNLOAD_PATH)

    # Load Excel file
    df = pd.read_excel(DOWNLOAD_PATH)
    logging.info(f"Excel Preview (first 8 rows):\n{df.head(8)}")

    # Extract columns assumed as 'Name', 'Email', 'Expiry Date'
    if not all(col in df.columns for col in ['Name', 'Email', 'Expiry Date']):
        logging.error("Required columns 'Name', 'Email', or 'Expiry Date' are missing from Excel.")
        return

    # Process agreements to find which are expiring today or soon
    df['Expiry Date'] = pd.to_datetime(df['Expiry Date'])
    today = pd.Timestamp.now().normalize()

    expiring_today = df[df['Expiry Date'] == today]

    logging.info(f"{len(expiring_today)} agreement(s) expiring today.")

    for _, row in expiring_today.iterrows():
        subject = f"Renewal Reminder for {row['Name']}"
        body = f"Dear {row['Name']},\n\nYour agreement expires today ({row['Expiry Date'].date()}). Please take necessary action."
        send_email(subject, body, [row['Email']])

if __name__ == '__main__':
    main()









