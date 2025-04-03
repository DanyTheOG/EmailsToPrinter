import imaplib
import email
import os
import datetime
from io import BytesIO
import smtplib
from email.message import EmailMessage

import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.pdfbase.pdfmetrics import stringWidth

# Environment variables from GitHub secrets
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
PRINTER_EMAIL = os.environ.get("PRINTER_EMAIL")

SEARCH_PHRASE = "Daily Leads Report"  # Text to search for in email body
ATTACHMENT_EXT = ".xlsx"             # Excel attachment extension
TEMP_PDF = "report.pdf"              # Temporary PDF file name

def connect_imap():
    """Connect to Gmail via IMAP."""
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASSWORD)
    return mail

def search_emails(mail):
    """Search the inbox for emails since yesterday."""
    mail.select("inbox")
    since_date = (datetime.date.today() - datetime.timedelta(days=1)).strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(SINCE "{since_date}")')
    if status != "OK":
        print("No messages found!")
        return []
    email_ids = messages[0].split()
    return email_ids

def is_recent(date_str):
    """Return True if the email date is within the last 12 hours."""
    try:
        email_date = email.utils.parsedate_to_datetime(date_str)
        now = datetime.datetime.now(email_date.tzinfo)
        delta = now - email_date
        return delta.total_seconds() <= 12 * 3600
    except Exception as e:
        print("Error parsing date:", e)
        return False

def get_email_body(msg):
    """Extract and return the plain text body from the email message."""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    return part.get_payload(decode=True).decode(part.get_content_charset() or 'utf-8')
                except Exception as e:
                    print("Error decoding email part:", e)
                    return ""
    else:
        try:
            return msg.get_payload(decode=True).decode(msg.get_content_charset() or 'utf-8')
        except Exception as e:
            print("Error decoding email:", e)
            return ""
    return ""

def get_latest_attachment(mail, email_ids):
    """
    Iterate through emails, filter by date and body content, and 
    return the attachment data from the most recent qualifying email.
    """
    latest_date = None
    attachment_data = None

    for eid in email_ids:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        date_hdr = msg.get("Date")
        if not date_hdr or not is_recent(date_hdr):
            continue

        # Check if the email body contains the search phrase
        body = get_email_body(msg)
        if SEARCH_PHRASE not in body:
            continue

        # Look for an Excel attachment in the email
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if filename and filename.lower().endswith(ATTACHMENT_EXT):
                    email_date = email.utils.parsedate_to_datetime(date_hdr)
                    if (latest_date is None) or (email_date > latest_date):
                        latest_date = email_date
                        attachment_data = part.get_payload(decode=True)
    return attachment_data

def convert_excel_to_pdf(excel_data):
    """Convert the first sheet of the Excel file to a PDF."""
    wb = openpyxl.load_workbook(BytesIO(excel_data), data_only=True)
    ws = wb.active  # Use the first sheet

    # Extract worksheet data into a list of lists
    data = []
    for row in ws.iter_rows(values_only=True):
        data.append([str(cell) if cell is not None else "" for cell in row])
    
    # Calculate approximate column widths based on max text width per column
    col_widths = []
    for col in zip(*data):
        max_width = max([stringWidth(str(item), 'Helvetica', 10) for item in col] + [0])
        col_widths.append(max_width + 10)  # add padding

    # Create PDF document in landscape orientation
    doc = SimpleDocTemplate(
        TEMP_PDF,
        pagesize=landscape(letter),
        rightMargin=30, leftMargin=30,
        topMargin=30, bottomMargin=18,
    )
    
    # Create a table with gridlines and centered text
    table = Table(data, colWidths=col_widths)
    style = TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ])
    table.setStyle(style)
    
    elements = [table]
    doc.build(elements)

def send_email(pdf_path):
    """Send an email with the PDF attached."""
    msg = EmailMessage()
    now = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    msg["Subject"] = f"Daily Report {now}"
    msg["From"] = GMAIL_USER
    msg["To"] = PRINTER_EMAIL
    msg.set_content("this email it to be printed")
    
    with open(pdf_path, "rb") as f:
        pdf_data = f.read()
    msg.add_attachment(pdf_data, maintype="application", subtype="pdf", filename="report.pdf")

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

        attachment_data = get_latest_attachment(mail, email_ids)
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
