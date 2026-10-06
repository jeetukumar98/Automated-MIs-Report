import glob
import os
import smtplib
from datetime import date
from email.message import EmailMessage

from dotenv import load_dotenv

load_dotenv()   # .env file se EMAIL_USER, EMAIL_PASS, EMAIL_TO padhta hai

EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASS = (os.getenv("EMAIL_PASS") or "").replace(" ", "")
EMAIL_TO = os.getenv("EMAIL_TO")


def latest_report(folder="output"):
    """output folder ki sabse nayi MIS report ka path do."""
    files = glob.glob(f"{folder}/MIS_Report_*.xlsx")
    if not files:
        raise FileNotFoundError("output folder me koi report nahi mili. Pehle mis_report.py chalao.")
    return max(files, key=os.path.getmtime)


def build_message(path):
    msg = EmailMessage()
    msg["Subject"] = f"Daily MIS Report - {date.today():%d-%b-%Y}"
    msg["From"] = EMAIL_USER
    msg["To"] = EMAIL_TO
    msg.set_content(
        "Hello,\n\n"
        "Please find attached today's MIS report.\n\n"
        "Regards,\nJeetu Kumar"
    )
    with open(path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(path),
        )
    return msg


def send(msg):
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(EMAIL_USER, EMAIL_PASS)
        server.send_message(msg)


def main():
    if not (EMAIL_USER and EMAIL_PASS and EMAIL_TO):
        raise SystemExit(".env me EMAIL_USER, EMAIL_PASS, EMAIL_TO teeno daalo.")
    path = latest_report()
    send(build_message(path))
    print("Email sent with:", os.path.basename(path))


if __name__ == "__main__":
    main()