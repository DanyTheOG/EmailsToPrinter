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
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
from reportlab.pdfbase.pdfmetrics import stringWidth

from zoneinfo import ZoneInfo  # Python 3.9+ for time zone support

# Environment variables from GitHub secrets
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_PASSWORD = os.environ.get("GMAIL_PASSWORD")
PRINTER_EMAIL = os.environ.get("PRINTER_EMAIL")

SEARCH_PHRASE = "Daily Leads Report"  # Must match the email content exactly
ATTACHMENT_EXT = ".xlsx"             # Excel attachment extension
TEMP_PDF = "report.pdf"              # Temporary PDF file name

def connect_imap():
    """Connect to Gmail via IMAP."""
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(GMAIL_USER, GMAIL_PASSWORD)
    return mail

def get_time_window():
    """
    Calculate the start and end time (in Madrid time) for processing emails.
    For Tue–Fri: window from yesterday 08:00 to today 09:00.
    For Monday: window from Friday 08:00 to Monday 09:00.
    """
    madrid_tz = ZoneInfo("Europe/Madrid")
    now = datetime.datetime.now(madrid_tz)
    # Ensure we only run on weekdays (0=Monday, 1=Tuesday, ..., 4=Friday)
    if now.weekday() > 4:
        print("Script should only run on weekdays. Exiting.")
        exit(0)
    if now.weekday() == 0:
        # Monday: window from Friday 08:00 to Monday 09:00 (3 days back)
        start_date = (now - datetime.timedelta(days=3)).date()
        start_time = datetime.datetime.combine(start_date, datetime.time(8, 0, tzinfo=madrid_tz))
    else:
        # Tue–Fri: window from yesterday 08:00 to today 09:00
        start_date = (now - datetime.timedelta(days=1)).date()
        start_time = datetime.datetime.combine(start_date, datetime.time(8, 0, tzinfo=madrid_tz))
    end_time = datetime.datetime.combine(now.date(), datetime.time(9, 0, tzinfo=madrid_tz))
    return start_time, end_time

def search_emails(mail, start_time):
    """
    Search the inbox for emails since the start date.
    The IMAP 'SINCE' command only uses the date part so extra filtering is done later.
    """
    mail.select("inbox")
    since_date_str = start_time.strftime("%d-%b-%Y")
    status, messages = mail.search(None, f'(SINCE "{since_date_str}")')
    if status != "OK":
        print("No messages found!")
        return []
    email_ids = messages[0].split()
    return email_ids

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

def get_attachments(mail, email_ids, start_time, end_time):
    """
    Iterate through emails and return a list of Excel attachment data from emails
    whose Date header (converted to Madrid time) falls within the window AND
    whose body contains the search phrase.
    """
    attachments = []
    madrid_tz = ZoneInfo("Europe/Madrid")
    for eid in email_ids:
        status, msg_data = mail.fetch(eid, "(RFC822)")
        if status != "OK":
            continue
        msg = email.message_from_bytes(msg_data[0][1])
        date_hdr = msg.get("Date")
        if not date_hdr:
            continue
        try:
            email_date = email.utils.parsedate_to_datetime(date_hdr)
        except Exception as e:
            print("Error parsing date:", e)
            continue
        # Convert email date to Madrid time
        email_date_madrid = email_date.astimezone(madrid_tz)
        if not (start_time <= email_date_madrid < end_time):
            continue
        # Check if the email body contains the search phrase
        body = get_email_body(msg)
        if SEARCH_PHRASE not in body:
            continue
        # Look for an Excel attachment
        for part in msg.walk():
            if part.get_content_disposition() == "attachment":
                filename = part.get_filename()
                if filename and filename.lower().endswith(ATTACHMENT_EXT):
                    attachments.append(part.get_payload(decode=True))
    return attachments

def convert_multiple_excels_to_pdf(excel_data_list):
    """
    Convert each Excel attachment (first sheet) to a table and combine them into a single PDF.
    Each table is separated by a page break.
    """
    styles =  getSampleStyleSheet = __import__('reportlab.lib.styles', fromlist=['getSampleStyleSheet']).getSampleStyleSheet()
    normal_style = styles["Normal"]

    elements = []
    for idx, excel_data in enumerate(excel_data_list):
        wb = openpyxl.load_workbook(BytesIO(excel_data), data_only=True)
        ws = wb.active

        # Extract worksheet data into a list of lists (as plain text)
        raw_data = []
        for row in ws.iter_rows(values_only=True):
            raw_data.append([str(cell) if cell is not None else "" for cell in row])
        
        # Wrap each cell's text in a Paragraph for automatic text wrapping
        data = []
        for row in raw_data:
            new_row = []
            for cell in row:
                new_row.append(Paragraph(cell, normal_style))
            data.append(new_row)
        
        # Calculate column widths based on plain text, capped by available width.
        col_widths = []
        if raw_data:
            num_cols = len(raw_data[0])
        else:
            num_cols = 0
        available_width = 792 - 60  # Landscape letter width minus left/right margins
        max_col_width = available_width / num_cols if num_cols > 0 else 100

        # For each column, calculate width
        for col in zip(*raw_data):
            calculated_width = max([stringWidth(str(item), 'Helvetica', 10) for item in col] + [0]) + 10
            col_widths.append(min(calculated_width, max_col_width))
        
        table = Table(data, colWidths=col_widths)
        style = TableStyle([
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ])
        table.setStyle(style)
        elements.append(table)
        # Add a page break between tables except after the last one
        if idx < len(excel_data_list) - 1:
            elements.append(PageBreak())
    
    doc = SimpleDocTemplate(
        TEMP_PDF,
        pagesize=landscape(letter),
        rightMargin=30, leftMargin=30,
        topMargin=30, bottomMargin=18,
    )
    doc.build(elements)

def send_email(pdf_path):
    """Send an email with the PDF attached."""
    msg = EmailMessage()
    now = datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    msg["Subject"] = f"Daily Report {now}"
    msg["From"] = GMAIL_USER
    msg["To"] = PRINTER_EMAIL
    msg.set_content("This email is to be printed.")
    
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
        madrid_tz = ZoneInfo("Europe/Madrid")
        now = datetime.datetime.now(madrid_tz)
        # Only run on weekdays
        if now.weekday() > 4:
            print("Script should only run on weekdays. Exiting.")
            return

        start_time, end_time = get_time_window()
        mail = connect_imap()
        email_ids = search_emails(mail, start_time)
        if not email_ids:
            print("No emails found matching the criteria.")
            return

        attachment_list = get_attachments(mail, email_ids, start_time, end_time)
        if not attachment_list:
            print("No valid Excel attachments found in the specified time window.")
            return

        print(f"Found {len(attachment_list)} attachment(s). Converting to PDF...")
        convert_multiple_excels_to_pdf(attachment_list)
        print("PDF generated successfully.")
        print("Sending email with PDF attached...")
        send_email(TEMP_PDF)
    except Exception as e:
        print("An error occurred:", e)
        raise

if __name__ == "__main__":
    main()
