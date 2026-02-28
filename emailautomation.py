#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Email Automation for Shortlisted & Non‑Shortlisted Candidates
Reads Shortlisted.xlsx and Non_Shortlisted.xlsx, sends appropriate emails via Gmail SMTP.
"""

import os
import time
import logging
import smtplib
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv

load_dotenv()

# ============================================
# CONFIGURATION FROM ENVIRONMENT
# ============================================
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")
SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", 587))

SHORTLISTED_FILE = os.getenv("SHORTLISTED_FILE")
NON_SHORTLISTED_FILE = os.getenv("NON_SHORTLISTED_FILE")

EMAIL_DELAY = int(os.getenv("EMAIL_DELAY", 2))
FILTER_BY_RANK = os.getenv("FILTER_BY_RANK", "True").lower() == "true"
MAX_RANK = int(os.getenv("MAX_RANK", 10))
RANK_COLUMN = "Rank"

# ============================================
# LOGGING
# ============================================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('email_campaign.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================
# EMAIL TEMPLATES
# ============================================
SHORTLISTED_TEMPLATE = """
<html>
<body>
    <h2>Congratulations {candidate_name}!</h2>
    <p>We are pleased to inform you that you have been <strong>shortlisted</strong> for the position of <strong>{job_title}</strong>.</p>
    <p>Your application stood out among many qualified candidates.</p>
    <h3>Next Steps</h3>
    <p>Our recruitment team will contact you within 48 hours to schedule an interview. Please keep an eye on your phone and email.</p>
    <p>Thank you for your interest in joining our team.</p>
    <p>Best regards,<br>
    Recruitment Team<br>
    Praval Info Tech PVT</p>
</body>
</html>
"""

NON_SHORTLISTED_TEMPLATE = """
<html>
<body>
    <h2>Update on your application</h2>
    <p>Dear {candidate_name},</p>
    <p>Thank you for applying for the <strong>{job_title}</strong> position at Your Company Name. We received a large number of applications and after careful review, we have decided to move forward with candidates whose experience more closely matches our current requirements.</p>
    <p>We encourage you to apply for future openings that match your profile. We wish you the best in your job search.</p>
    <p>Sincerely,<br>
    Recruitment Team<br>
    Praval Info Tech PVT</p>
</body>
</html>
"""

# ============================================
# FUNCTIONS
# ============================================
def send_email(recipient, subject, html_body):
    msg = MIMEMultipart("alternative")
    msg["From"] = SENDER_EMAIL
    msg["To"] = recipient
    msg["Subject"] = subject
    part = MIMEText(html_body, "html")
    msg.attach(part)

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.send_message(msg)
        logger.info(f"Email sent to {recipient}")
        return True
    except Exception as e:
        logger.error(f"Failed to send to {recipient}: {e}")
        return False

def process_file(file_path, template, status_label, filter_by_rank=False):
    if not os.path.exists(file_path):
        logger.warning(f"File not found: {file_path}")
        return 0

    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        df.columns = [str(col).strip().lower() for col in df.columns]
    except Exception as e:
        logger.error(f"Error reading {file_path}: {e}")
        return 0

    name_col = next((c for c in df.columns if 'name' in c), None)
    email_col = next((c for c in df.columns if 'email' in c), None)
    job_col = next((c for c in df.columns if 'job' in c or 'position' in c), None)
    rank_col = None
    if filter_by_rank:
        rank_col = next((c for c in df.columns if 'rank' in c), None)
        if not rank_col:
            logger.warning("No rank column found, skipping rank filter.")
            filter_by_rank = False

    if not name_col or not email_col:
        logger.error(f"Missing name or email column in {file_path}")
        return 0
    if not job_col:
        logger.warning("No job column found – emails will omit job title.")

    if filter_by_rank:
        df = df[df[rank_col] <= MAX_RANK].copy()
        logger.info(f"Filtered to {len(df)} candidates with rank ≤ {MAX_RANK}.")

    sent = 0
    for _, row in df.iterrows():
        name = row[name_col]
        email = row[email_col]
        job = row[job_col] if job_col else "the position"

        if pd.isna(email) or str(email).strip() == "" or str(email).lower() == "not found":
            logger.warning(f"No valid email for {name}, skipping.")
            continue

        subject = f"Congratulations on being shortlisted for {job}!" if status_label == "Shortlisted" else f"Update on your application for {job}"
        html = template.format(candidate_name=name, job_title=job)

        if send_email(email, subject, html):
            sent += 1
        time.sleep(EMAIL_DELAY)

    logger.info(f"{file_path}: Sent {sent} {status_label} emails.")
    return sent

def main():
    print("\n========================================")
    print("RECRUITMENT EMAIL AUTOMATION")
    print("========================================\n")

    if len(SENDER_PASSWORD) != 16:
        logger.warning("App password should be 16 characters. Please check your configuration.")

    total = 0
    if os.path.exists(SHORTLISTED_FILE):
        total += len(pd.read_excel(SHORTLISTED_FILE, engine='openpyxl'))
    if os.path.exists(NON_SHORTLISTED_FILE):
        total += len(pd.read_excel(NON_SHORTLISTED_FILE, engine='openpyxl'))

    if total == 0:
        logger.warning("No candidate files found.")
        return

    logger.info(f"About to send up to {total} emails (delay {EMAIL_DELAY}s each).")
    # Automatically proceed without asking for confirmation
    process_file(SHORTLISTED_FILE, SHORTLISTED_TEMPLATE, "Shortlisted", filter_by_rank=FILTER_BY_RANK)
    process_file(NON_SHORTLISTED_FILE, NON_SHORTLISTED_TEMPLATE, "Non Shortlisted", filter_by_rank=False)

    logger.info("Email campaign completed.")

if __name__ == "__main__":
    main()