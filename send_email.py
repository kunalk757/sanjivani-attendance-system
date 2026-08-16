import os
import csv
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

RECIPIENT_EMAIL = "rohitvinchu7754@gmail.com"
STUDENT_FOLDER = "students"
ATTENDANCE_CSV = "attendance.csv"
ATTENDANCE_XLSX = "attendance.xlsx"


def get_email_credentials():
    """Retrieve sender email and password from environment variables."""
    sender_email = (
        os.environ.get("SENDER_EMAIL")
        or os.environ.get("EMAIL_USER")
        or os.environ.get("GMAIL_USER")
    )
    sender_password = (
        os.environ.get("SENDER_PASSWORD")
        or os.environ.get("EMAIL_PASS")
        or os.environ.get("APP_PASSWORD")
        or os.environ.get("GMAIL_PASS")
    )
    return sender_email, sender_password


def send_email_report(sender_email=None, sender_password=None):
    """
    Send attendance report email with attendance.xlsx attached to rohitvinchu7754@gmail.com.

    Returns:
        (success: bool, message: str)
    """
    if not sender_email or not sender_password:
        env_email, env_pass = get_email_credentials()
        sender_email = sender_email or env_email
        sender_password = sender_password or env_pass

    if not sender_email or not sender_password:
        return False, "SENDER_EMAIL and SENDER_PASSWORD environment variables are not set."

    now = datetime.now()
    today_str = now.strftime("%Y-%m-%d")
    date_time_str = now.strftime("%Y-%m-%d %I:%M:%S %p")

    # Calculate student statistics
    total_students = 0
    if os.path.exists(STUDENT_FOLDER):
        total_students = len([d for d in os.listdir(STUDENT_FOLDER) if os.path.isdir(os.path.join(STUDENT_FOLDER, d))])

    present_count = 0
    if os.path.exists(ATTENDANCE_CSV):
        try:
            with open(ATTENDANCE_CSV, mode="r", newline="", encoding="utf-8") as f:
                reader = csv.reader(f)
                next(reader, None)  # Skip header
                present_count = sum(1 for row in reader if row)
        except Exception:
            pass

    absent_count = max(0, total_students - present_count)

    # Compose MIMEMultipart email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = f"Sanjivani Attendance Report - {today_str}"

    body_text = f"""Dear Admin,

The Sanjivani Attendance System session has been completed successfully.

--------------------------------------------------
ATTENDANCE REPORT SUMMARY
--------------------------------------------------
Attendance Session Completed : Yes
Date & Time                  : {date_time_str}
Total Registered Students    : {total_students}
Present Student Count        : {present_count}
Absent Student Count         : {absent_count}
--------------------------------------------------

Confirmation: The complete Excel attendance report (attendance.xlsx) is attached to this email.

Best Regards,
Sanjivani Attendance System
AI-Powered Face Recognition & Smart Attendance
"""

    msg.attach(MIMEText(body_text, 'plain'))

    # Attach attendance.xlsx file
    if os.path.exists(ATTENDANCE_XLSX) and os.path.getsize(ATTENDANCE_XLSX) > 0:
        try:
            with open(ATTENDANCE_XLSX, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename=attendance.xlsx",
            )
            msg.attach(part)
        except Exception as e:
            return False, f"Could not attach attendance.xlsx: {e}"
    else:
        return False, "attendance.xlsx file does not exist or is empty."

    # Connect to SMTP server (Gmail 587 TLS)
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
        server.quit()
        return True, f"✓ Attendance report email successfully sent to {RECIPIENT_EMAIL}!"
    except smtplib.SMTPAuthenticationError:
        return False, "SMTP Authentication Failed: Incorrect email or Google App Password."
    except Exception as e:
        return False, f"SMTP Error: {e}"
