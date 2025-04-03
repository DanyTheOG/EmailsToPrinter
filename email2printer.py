import imaplib
import email
from email.header import decode_header
import os
import datetime
from io import BytesIO
import smtplib
from email.message import EmailMessage

import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import stringWidth


# Environment variables from GitHub secrets
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
PRINTER_EMAIL = os.environ.get("PRINTER_EMAIL")

# Constants
SEARCH_PHRASE = "Daily Lead Report"
ATTACHMENT_EXT = ".xlsx"
TEMP_PDF = "report.pdf"

def connect_imap():
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASSWORD)
    return mail

def search_emails(mail):
    mail.select("inbox")
    # Search for emails from self containing the phrase in the body
    status, messages = mail.search(None, f'(FROM "{GMAIL_USER}" BODY "{SEARCH_PHRASE}")')
    if status != "OK":
        print("No messages found!")
        return []
    # messages is a byte string of space-separated email IDs
    email_ids = messages[0].split()
    return email_ids

def is_recent(date_str):
    """
    Check if the email date (from the header) is within the last 12 hours.
    The date_str is typically in RFC 2822 format.
    """
    try:
        # Parse email date
        email_date = email.utils.parsedate_to_datetime(date_str)
        now = datetime.datetime.now(email_date.tzinfo)
        delta = now - email_date
        return delta.total_seconds() <= 12 * 3600
    except Exception as e:
        print("Error parsing date:", e)
        return False

def get_latest_attachment(mail, email_ids):
    latest_email = None
    latest_date = None
    attachment_data = None

    for eid in email_ids:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        # Check the Date header
        date_hdr = msg.get("Date")
        if not date_hdr or not is_recent(date_hdr):
            continue

        # Process each part looking for an Excel attachment
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if filename and filename.lower().endswith(ATTACHMENT_EXT):
                    # Get the email date as a datetime object
                    email_date = email.utils.parsedate_to_datetime(date_hdr)
                    if (latest_date is None) or (email_date > latest_date):
                        latest_date = email_date
                        latest_email = msg
                        attachment_data = part.get_payload(decode=True)
    return latest_email, attachment_data

def convert_excel_to_pdf(excel_data):
    # Load workbook from bytes
    wb = openpyxl.load_workbook(BytesIO(excel_data), data_only=True)
    ws = wb.active  # use the first sheet

    # Gather data from worksheet into a list of lists
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append([str(cell) if cell is not None else "" for cell in row])

    # Calculate approximate column widths based on maximum text width in each column
    col_widths = []
    for col in zip(*data):
        max_width = max([stringWidth(str(item), 'Helvetica', 10) for item in col] + [0])
        # add some padding
        col_widths.append(max_width + 10)

    # Create PDF document in landscape orientation
    doc = SimpleDocTemplate(
        TEMP_PDF,
        pagesize=landscape(letter),
        rightMargin=30, leftMargin=30,
        topMargin=30, bottomMargin=18,
    )

    # Create a table with the data
    table = Table(data, colWidths=col_widths)
    # Apply table style for gridlines and alignment
    style = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ])
    table.setStyle(style)

    # Build PDF
    elements = [table]
    doc.build(elements)

def send_email(pdf_path):
    # Compose email
    msg = EmailMessage()
    now = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    msg["Subject"] = f"Daily Report {now}"
    msg["From"] = GMAIL_USER
    msg["To"] = PRINTER_EMAIL
    msg.set_content("this email it to be printed")

    # Attach PDF
    with open(pdf_path, "rb") as f:
        pdf_data = f.read()
    msg.add_attachment(pdf_data, maintype="application", subtype="pdf", filename="report.pdf")

    # Send email via Gmail SMTP
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print("Email sent successfully.")
    except Exception as e:
        print("Error sending email:", e)
        raise

def main():
    try:
        mail = connect_imap()
        email_ids = search_emails(mail)
        if not email_ids:
            print("No emails found matching the criteria.")
            return

        _, attachment_data = get_latest_attachment(mail, email_ids)
        if not attachment_data:
            print("No valid Excel attachment found in the recent emails.")
            return

        print("Excel attachment found. Converting to PDF...")
        convert_excel_to_pdf(attachment_data)
        print("PDF generated successfully.")

        print("Sending email with PDF attached...")
        send_email(TEMP_PDF)
    except Exception as e:
        print("An error occurred:", e)
        raise

if __name__ == "__main__":
    main()
