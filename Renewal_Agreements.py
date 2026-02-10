import os
import gdown
import pandas as pd
from datetime import datetime, date
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
import logging

# ✅ GitHub Secrets Mapping
SENDER_EMAIL = os.getenv("SMTP_USER")
APP_PASSWORD = os.getenv("SMTP_PASS")
FILE_ID = os.getenv("GDRIVE_FILE_ID")

RECIPIENTS = [os.getenv("RECIPIENT_DEFAULT", "")]
DOWNLOAD_PATH = Path(os.getenv("DOWNLOAD_PATH", "Renewal.xlsx"))

SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

SEND_CONFIRMATION = False   
LOG_FILE = Path(os.getenv("LOG_FILE", "renewal.log"))

TEMPLATE_SUBJECT = "Service Renewal Alert - Operations"

logging.basicConfig( 
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

TEMPLATE_BODY = """
<p>Hi {Team_Member_Name},</p>

<p>
This is to inform you that the <b>{Service_Name}</b> service for (<b>Business_Name</b>) is coming up for renewal on <b>{Service_End_Date}</b>.
</p>

<p>
Please make sure all deliverables are delivered. If not, reach out to Famida and take the necessary action.
</p>

<p>
Pause services on the mentioned date and wait for Famida's heads-up regarding the renewal to continue services.
</p>

<p>
Kindly update the service status to this mail on <b>{Service_End_Date}</b> to avoid further reminders.
</p>

<p>
Thanks,<br>
NTQS Digital<br>
Operations Team
</p>
"""

def download_from_drive(file_id: str, dest: Path):
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    logging.info(f"Downloading {url} -> {dest} ...")
    try:
        gdown.download(url, str(dest), quiet=False)
        logging.info("Download complete.")
    except Exception:
        logging.exception("Failed to download file from Google Drive.")
        raise

def detect_header_and_load(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")

    header_idx = None
    search_tokens = {"expiry", "expiry date", "email", "name", "file", "due", "end", "expires"}

    max_scan = min(50, len(raw))

    for i in range(max_scan):
        row_vals = [str(x).strip().lower() for x in raw.iloc[i].fillna("")]
        combined = " ".join(row_vals)

        if any(tok in combined for tok in search_tokens):
            header_idx = i
            logging.info(f"Detected header at row {header_idx}")
            break

    if header_idx is None:
        header_idx = 0

    df = pd.read_excel(path, header=header_idx, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    return df

def build_message(sender: str, to_list: list, subject: str, body: str) -> EmailMessage:
    msg = EmailMessage()

    msg['From'] = sender
    msg['To'] = ", ".join(to_list)
    msg['Subject'] = subject

    msg.set_content("HTML Email Required")
    msg.add_alternative(body, subtype='html')

    return msg

def send_email(msg: EmailMessage):
    context = ssl.create_default_context()

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
            smtp.ehlo()
            smtp.starttls(context=context)
            smtp.login(SENDER_EMAIL, APP_PASSWORD)
            smtp.send_message(msg)

        logging.info(f"Email sent to {msg['To']}")

    except Exception:
        logging.exception("Email sending failed")
        raise

def make_agreements_list(df: pd.DataFrame) -> list:
    agreements = []

    expiry_cols = [c for c in df.columns if 'expiry' in c.lower() or 'end' in c.lower()]
    email_cols = [c for c in df.columns if 'email' in c.lower()]
    name_cols = [c for c in df.columns if 'name' in c.lower()]
    service_cols = [c for c in df.columns if 'service' in c.lower()]
    business_cols = [c for c in df.columns if 'business' in c.lower()]

    expiry_col = expiry_cols[0]
    email_col = email_cols[0]
    name_col = name_cols[0] if name_cols else None
    service_col = service_cols[0] if service_cols else None
    business_col = business_cols[0] if business_cols else None

    for _, row in df.iterrows():

        expiry_dt = pd.to_datetime(row.get(expiry_col), errors='coerce', dayfirst=True)
        if pd.isna(expiry_dt):
            continue

        agreements.append({
            'expiry_date': expiry_dt,
            'email': str(row.get(email_col, '')).strip(),
            'name': str(row.get(name_col, '')).strip() if name_col else "",
            'service': str(row.get(service_col, '')).strip() if service_col else "",
            'business': str(row.get(business_col, '')).strip() if business_col else ""
        })

    return agreements

def send_renewal_reminder(agreement: dict):
    team_member = agreement.get('name') or agreement.get('email') or "Team"

    body = TEMPLATE_BODY.format(
        Team_Member_Name=team_member,
        Service_Name=agreement.get('service'),
        Business_Name=agreement.get('business'),
        Service_End_Date=agreement['expiry_date'].strftime('%d-%m-%Y')
    )

    msg = build_message(
        SENDER_EMAIL,
        [agreement.get('email')],
        TEMPLATE_SUBJECT,
        body
    )

    send_email(msg)

def run_reminders_and_alerts(agreements: list):

    today = date.today()

    for agreement in agreements:

        days_left = (agreement['expiry_date'].date() - today).days

        if 1 <= days_left <= 5:
            send_renewal_reminder(agreement)

        if days_left == 0:
            send_renewal_reminder(agreement)

def main_run_once():

    download_from_drive(FILE_ID, DOWNLOAD_PATH)

    df = detect_header_and_load(DOWNLOAD_PATH)

    agreements = make_agreements_list(df)

    run_reminders_and_alerts(agreements)

if __name__ == "__main__":
    main_run_once()









