# =========================================================
# HireFlow AI - HR + Company Email Automation
# Sends emails to BOTH:
#   1) Company emails from `Emails` column
#   2) HR emails from `"hr@" + Domain` and `"careers@" + Domain` columns
# Supports multiple emails in one cell separated by comma, newline, semicolon, or space.
# Tracks SENT / PARTIAL / FAILED / NO EMAIL in Excel.
# =========================================================

import os
import re
import time
from datetime import datetime

import pandas as pd
from dotenv import load_dotenv

from email_sender import send_email
from email_templates import backend_email, ml_ds_email

load_dotenv()

# =========================================================
# CONFIG
# =========================================================

FILE_NAME = "hr_email.xlsx"

TEST_MODE = False
TEST_EMAIL = os.getenv("TEST_EMAIL")

BACKEND_RESUME = "resumes/Muhammad_Zaid.pdf"
ML_DS_RESUME = "resumes/MuhammadZaid.pdf"

MAX_COMPANIES_PER_RUN = 5
DELAY_SECONDS = 5

# =========================================================
# COLUMN NAMES — match your actual Excel headers exactly
# =========================================================

# BUG FIX 1: Excel mein "hr_email" column nahi tha.
# Actual columns hain: "Emails", '"hr@" + Domain', '"careers@" + Domain'
COL_COMPANY_EMAIL = "Emails"
COL_HR_EMAIL      = '"hr@" + Domain'
COL_CAREERS_EMAIL = '"careers@" + Domain'

# =========================================================
# EMAIL HELPERS
# =========================================================

EMAIL_PATTERN = r"[\w\.-]+@[\w\.-]+\.\w+"


def extract_emails(value):
    """Return all valid emails from a cell."""
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    emails = re.findall(EMAIL_PATTERN, text)
    cleaned = []
    seen = set()
    for email in emails:
        email = email.strip().lower()
        if email and email not in seen:
            cleaned.append(email)
            seen.add(email)
    return cleaned


def extract_company_name(row, all_emails):
    """Prefer Company column, otherwise extract from email domain."""
    company_value = row.get("Company", "")
    if not pd.isna(company_value) and str(company_value).strip():
        company = str(company_value).strip()
        company = company.replace("www.", "").split(".")[0]
        return company.capitalize()
    if all_emails:
        domain = all_emails[0].split("@")[-1]
        return domain.split(".")[0].capitalize()
    return "Company"


# =========================================================
# SAFE EMAIL SENDER
# =========================================================


def safe_send(to_email, subject, body, email_type, attachment_path=None):
    try:
        if TEST_MODE:
            final_email = TEST_EMAIL
            if not final_email:
                print("❌ TEST_EMAIL missing in .env file")
                return False
            print(f"\n🧪 TEST MODE | Original: {to_email} → Sending to: {TEST_EMAIL}")
        else:
            final_email = to_email
            print(f"\n🚀 Sending To: {final_email}")

        send_email(final_email, subject, body, attachment_path)
        print(f"✅ Sent: {email_type}")
        if attachment_path:
            print(f"📎 Attachment: {attachment_path}")
        return True

    except Exception as e:
        print(f"❌ Failed: {email_type} | Error: {e}")
        return False


# =========================================================
# SEND BOTH POSITIONS TO ONE EMAIL ADDRESS
# BUG FIX 2: Helper so we don't duplicate this logic
# =========================================================


def send_both_positions(email, company, email_label, results):
    """Send Backend + ML/DS email to one address. Updates results dict in-place."""

    # --- Backend position ---
    subject, body = backend_email(company)
    print(f"\n📧 {email_label} — BACKEND ROLE → {email}")
    print(f"   Subject: {subject}")
    results["total"] += 1
    if safe_send(email, subject, body, f"{email_label} Backend", BACKEND_RESUME):
        results["success"] += 1
    else:
        results["errors"].append(f"{email_label} Backend failed: {email}")
    time.sleep(DELAY_SECONDS)

    # --- ML / DS position ---
    subject, body = ml_ds_email(company)
    print(f"\n📧 {email_label} — ML/DS ROLE → {email}")
    print(f"   Subject: {subject}")
    results["total"] += 1
    if safe_send(email, subject, body, f"{email_label} ML/DS", ML_DS_RESUME):
        results["success"] += 1
    else:
        results["errors"].append(f"{email_label} ML/DS failed: {email}")
    time.sleep(DELAY_SECONDS)


# =========================================================
# SAVE EXCEL
# =========================================================


def save_excel_file(df, file_name):
    try:
        df.to_excel(file_name, index=False)
        print("📄 Excel updated.")
        return True
    except PermissionError:
        backup = f"hr_email_updated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        df.to_excel(backup, index=False)
        print(f"❌ Excel open hai, please close karo. Backup: {backup}")
        return False
    except Exception as e:
        print(f"❌ Excel save error: {e}")
        return False


# =========================================================
# LOAD EXCEL
# =========================================================

try:
    df = pd.read_excel(FILE_NAME)
    df.columns = df.columns.str.strip()
    print(f"\n✅ {len(df)} records loaded.\n")
except Exception as e:
    print(f"❌ Excel load error:\n{e}")
    exit()

# =========================================================
# REQUIRED COLUMNS CHECK
# BUG FIX 3: Check for actual column names in your Excel
# =========================================================

required_columns = [COL_COMPANY_EMAIL, COL_HR_EMAIL, COL_CAREERS_EMAIL]
for col in required_columns:
    if col not in df.columns:
        print(f"❌ Missing column: '{col}'")
        print(f"   Available columns: {df.columns.tolist()}")
        exit()

# =========================================================
# STATUS COLUMNS
# =========================================================

for col in ["Status", "Sent_Date", "Last_Error", "Company_Email_Status", "HR_Email_Status"]:
    if col not in df.columns:
        df[col] = ""
    df[col] = df[col].astype(str).replace("nan", "")

# =========================================================
# PROCESSING
# =========================================================

companies_processed_this_run = 0

for index, row in df.iterrows():
    status = str(row.get("Status", "")).strip().upper()

    if status == "SENT":
        continue

    if MAX_COMPANIES_PER_RUN is not None and companies_processed_this_run >= MAX_COMPANIES_PER_RUN:
        break

    # --- Extract emails from all three columns ---
    company_emails = extract_emails(row.get(COL_COMPANY_EMAIL, ""))

    # BUG FIX 3: Combine hr@ + careers@ into one HR list, remove duplicates
    hr_emails_raw = (
        extract_emails(row.get(COL_HR_EMAIL, ""))
        + extract_emails(row.get(COL_CAREERS_EMAIL, ""))
    )
    seen_hr = set()
    hr_emails = []
    for e in hr_emails_raw:
        if e not in seen_hr:
            hr_emails.append(e)
            seen_hr.add(e)

    all_emails = company_emails + hr_emails
    company = extract_company_name(row, all_emails)

    print("=" * 90)
    print(f"📌 Company       : {company}")
    print(f"🔢 Excel Row     : {index + 2}")
    print(f"🏢 Company Emails: {company_emails}")
    print(f"👤 HR Emails     : {hr_emails}")
    print("=" * 90)

    if not company_emails and not hr_emails:
        print("⚠️ No valid email found.")
        df.at[index, "Status"] = "NO EMAIL"
        df.at[index, "Last_Error"] = "No valid email in any column"
        save_excel_file(df, FILE_NAME)
        continue

    hr_results      = {"total": 0, "success": 0, "errors": []}
    company_results = {"total": 0, "success": 0, "errors": []}

    # --- Send to HR emails (hr@ + careers@) ---
    for hr_email in hr_emails:
        send_both_positions(hr_email, company, "HR", hr_results)

    # --- Send to Company emails (skip duplicates already sent as HR) ---
    for company_email in company_emails:
        if company_email in hr_emails:
            print(f"\n⚠️ Duplicate skipped (already sent as HR): {company_email}")
            continue
        send_both_positions(company_email, company, "Company", company_results)

    # --- Status update ---
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    all_errors = hr_results["errors"] + company_results["errors"]
    total_attempts = hr_results["total"] + company_results["total"]
    total_success  = hr_results["success"] + company_results["success"]

    def status_label(results):
        if results["total"] == 0:
            return "NO EMAIL"
        if results["success"] == results["total"]:
            return "SENT"
        if results["success"] > 0:
            return "PARTIAL"
        return "FAILED"

    df.at[index, "HR_Email_Status"]      = status_label(hr_results)
    df.at[index, "Company_Email_Status"] = status_label(company_results)

    if total_attempts == 0:
        df.at[index, "Status"]     = "NO EMAIL"
        df.at[index, "Last_Error"] = "No valid non-duplicate email attempts"
    elif total_success == total_attempts:
        df.at[index, "Status"]     = "SENT"
        df.at[index, "Sent_Date"]  = now
        df.at[index, "Last_Error"] = ""
        companies_processed_this_run += 1
        print(f"\n✅ SENT: {company}")
    elif total_success > 0:
        df.at[index, "Status"]     = "PARTIAL"
        df.at[index, "Sent_Date"]  = now
        df.at[index, "Last_Error"] = ", ".join(all_errors)
        companies_processed_this_run += 1
        print(f"\n⚠️ PARTIAL: {company}")
    else:
        df.at[index, "Status"]     = "FAILED"
        df.at[index, "Last_Error"] = ", ".join(all_errors)
        print(f"\n❌ FAILED: {company}")

    save_excel_file(df, FILE_NAME)

# =========================================================
# FINAL SAVE
# =========================================================

save_excel_file(df, FILE_NAME)

print("=" * 90)
print("✅ Email process complete.")
print(f"📌 Companies processed this run: {companies_processed_this_run}")
print("=" * 90)