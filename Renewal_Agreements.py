import os
import gdown
import pandas as pd
from datetime import datetime
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
import logging
import json

# ---------------- ENV ---------------- #

SENDER_EMAIL = os.getenv("SMTP_USER")
APP_PASSWORD = os.getenv("SMTP_PASS")
FILE_ID = os.getenv("GDRIVE_FILE_ID")

DOWNLOAD_PATH = Path(os.getenv("DOWNLOAD_PATH", "Renewal.xlsx"))
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

LOG_FILE = Path(os.getenv("LOG_FILE", "renewal.log"))
HISTORY_FILE = Path("sent_history.json")

TEMPLATE_SUBJECT = "Service Renewal Alert - Operations"

# ---------------- LOG ---------------- #

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"), logging.StreamHandler()]
)

# ---------------- EMAIL TEMPLATE ---------------- #

TEMPLATE_BODY = """
<p>Hi {Team_Member_Name},</p>

<p>
This is to inform you that the <b>{Service_Name}</b> service for (<b>{Business_Name}</b>) is coming up for renewal on <b>{Service_End_Date}</b>.
</p>

<p>Please make sure all deliverables are delivered. If not, reach out to Famida and take the necessary action.</p>

<p>Pause services on the mentioned date and wait for Famida's heads-up regarding the renewal to continue services.</p>

<p>Kindly update the service status to this mail on <b>{Service_End_Date}</b> to avoid further reminders.</p>

<p>
Thanks,<br>
NTQS Digital<br>
Operations Team
</p>
"""

# ---------------- HELPERS ---------------- #

def clean(val):
    if pd.isna(val):
        return ""
    return str(val).strip()

def find_column(possible, columns):
    for key in possible:
        for col in columns:
            if key in col:
                return col
    return None

# ---------------- HISTORY ---------------- #

def load_history():
    if HISTORY_FILE.exists():
        with open(HISTORY_FILE, "r") as f:
            return json.load(f)
    return {}

def save_history(history):
    with open(HISTORY_FILE, "w") as f:
        json.dump(history, f, indent=2)

def already_sent(agreement, tag):
    history = load_history()
    key = f"{agreement['email']}_{agreement['expiry_date'].date()}_{tag}"
    return history.get(key, False)

def mark_sent(agreement, tag):
    history = load_history()
    key = f"{agreement['email']}_{agreement['expiry_date'].date()}_{tag}"
    history[key] = True
    save_history(history)

# ---------------- DOWNLOAD ---------------- #

def download_from_drive(file_id: str, dest: Path):
    url = f"https://drive.google.com/uc?id={file_id}&export=download"
    logging.info(f"Downloading sheet...")
    gdown.download(url, str(dest), quiet=False)

# ---------------- LOAD EXCEL ---------------- #

def detect_header_and_load(path: Path) -> pd.DataFrame:
    raw = pd.read_excel(path, header=None, engine="openpyxl")

    header_idx = 0
    search_words = ["email", "name", "expiry", "end", "service"]

    for i in range(min(50, len(raw))):
        row = " ".join(str(x).lower() for x in raw.iloc[i].fillna(""))
        if any(w in row for w in search_words):
            header_idx = i
            break

    df = pd.read_excel(path, header=header_idx, engine="openpyxl")

    # NORMALIZE HEADERS (IMPORTANT FIX)
    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r'\s+', ' ', regex=True)
    )

    df = df.dropna(how="all").reset_index(drop=True)

    logging.info(f"Detected Columns: {list(df.columns)}")
    return df

# ---------------- BUILD AGREEMENTS ---------------- #

def make_agreements_list(df: pd.DataFrame):

    expiry_col = find_column(['expiry', 'end date', 'renewal', 'due'], df.columns)
    email_col = find_column(['email'], df.columns)
    name_col = find_column(['name'], df.columns)
    service_col = find_column(['service'], df.columns)
    business_col = find_column(['business', 'client', 'company'], df.columns)

    logging.info(f"""
Column Mapping:
expiry -> {expiry_col}
email -> {email_col}
name -> {name_col}
service -> {service_col}
business -> {business_col}
""")

    agreements = []

    for _, row in df.iterrows():

        expiry_dt = pd.to_datetime(row.get(expiry_col), errors='coerce', dayfirst=True)
        if pd.isna(expiry_dt):
            continue

        agreements.append({
            'expiry_date': expiry_dt,
            'email': clean(row.get(email_col)),
            'name': clean(row.get(name_col)),
            'service': clean(row.get(service_col)),
            'business': clean(row.get(business_col)),
        })

    return agreements

# ---------------- EMAIL ---------------- #

def build_message(to_email, body):
    msg = EmailMessage()
    msg['From'] = SENDER_EMAIL
    msg['To'] = to_email
    msg['Subject'] = TEMPLATE_SUBJECT
    msg.set_content("HTML Email Required")
    msg.add_alternative(body, subtype='html')
    return msg

def send_email(msg):
    context = ssl.create_default_context()
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as smtp:
        smtp.starttls(context=context)
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)
    logging.info(f"Sent -> {msg['To']}")

def send_renewal_reminder(agreement):

    body = TEMPLATE_BODY.format(
        Team_Member_Name=agreement['name'] or "Team",
        Service_Name=agreement['service'],
        Business_Name=agreement['business'],
        Service_End_Date=agreement['expiry_date'].strftime('%d-%m-%Y')
    )

    msg = build_message(agreement['email'], body)
    send_email(msg)

# ---------------- SCHEDULER ---------------- #

def run_reminders_and_alerts(agreements):

    now = datetime.now()
    today = now.date()

    for agreement in agreements:

        days_left = (agreement['expiry_date'].date() - today).days

        if 1 <= days_left <= 5:
            tag = f"pre_{days_left}"
            if not already_sent(agreement, tag):
                send_renewal_reminder(agreement)
                mark_sent(agreement, tag)

        if days_left == 0:

            if now.hour == 9 and 25 <= now.minute <= 35:
                if not already_sent(agreement, "morning"):
                    send_renewal_reminder(agreement)
                    mark_sent(agreement, "morning")

            if now.hour == 17 and 25 <= now.minute <= 35:
                if not already_sent(agreement, "evening"):
                    send_renewal_reminder(agreement)
                    mark_sent(agreement, "evening")

# ---------------- MAIN ---------------- #

def main_run_once():
    download_from_drive(FILE_ID, DOWNLOAD_PATH)
    df = detect_header_and_load(DOWNLOAD_PATH)
    agreements = make_agreements_list(df)
    run_reminders_and_alerts(agreements)

if __name__ == "__main__":
    main_run_once()


