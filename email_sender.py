import smtplib
import os
import mimetypes
from email.message import EmailMessage
from dotenv import load_dotenv

load_dotenv()

SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")


def send_email(to_email, subject, body, attachment_path=None):
    if not SENDER_EMAIL or not APP_PASSWORD:
        raise ValueError("SENDER_EMAIL or APP_PASSWORD missing in .env file")

    msg = EmailMessage()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)

    if attachment_path:
        if not os.path.exists(attachment_path):
            raise FileNotFoundError(f"Attachment not found: {attachment_path}")

        mime_type, _ = mimetypes.guess_type(attachment_path)
        maintype, subtype = (
            mime_type.split("/") if mime_type else ("application", "octet-stream")
        )

        with open(attachment_path, "rb") as file:
            msg.add_attachment(
                file.read(),
                maintype=maintype,
                subtype=subtype,
                filename=os.path.basename(attachment_path),
            )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)
