import gspread
from oauth2client.service_account import ServiceAccountCredentials
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Setup for Google Sheets API
SCOPE = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
CREDS_FILE = "credentials.json"
SPREADSHEET_NAME = "Trip Bookings"

SENDER_EMAIL = "priya.project.trial@gmail.com"
SENDER_PASSWORD = "wgms fsny tiwh ehtb"  # App password
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Connect to Google Sheet
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDS_FILE, SCOPE)
client = gspread.authorize(creds)
sheet = client.open(SPREADSHEET_NAME).worksheet("Sheet1")
data = sheet.get_all_records()

# Helper to format supplier email
def generate_supplier_email(
    supplier_name,
    num_adults,
    num_children,
    accommodation_type,
    destination,
    check_in_date,
    check_out_date,
    num_nights
):
    subject = f"Trip Query for {destination} - {num_adults + num_children} PAX"

    # Determine adult phrasing
    adult_phrase = f"{num_adults} adult" if num_adults == 1 else f"{num_adults} adults"

    # Determine child phrasing only if applicable
    if num_children:
        child_phrase = f"{num_children} child" if num_children == 1 else f"{num_children} children"
        pax_line = f"{num_adults + num_children} PAX ({adult_phrase} and {child_phrase})"
    else:
        pax_line = f"{num_adults + num_children} PAX ({adult_phrase})"

    body = f"""
        Dear {supplier_name},

        I hope this message finds you well.

        We are planning a trip for a group of {pax_line} to {destination}.

        Here are the trip details:
        • Check-in Date: {check_in_date}
        • Check-out Date: {check_out_date}
        • Number of Nights: {num_nights}
        • Preferred Accommodation Type: {accommodation_type}
    """

    if num_children:
        body += "\nWould it be possible to add extra bedding for the children in the specified accommodation?"

    body += """

    We would appreciate it if you could share a quote and availability at your earliest convenience.

    Warm regards,  
    [Your Name]  
    [Your Company Name]  
    [Your Contact Info]
    """

    return subject, body

supplier_name = "ABC"
supplier_email = 'priya.keshrinis@gmail.com'

# Loop through each row and send email
for row in data:
    adults = int(row['Adults'])
    children = int(row['Children']) if row['Children'] else 0
    accommodation = row['Accommodation Type']
    destination = row['Destination']
    check_in = row['Checkin']
    check_out = row['Checkout']
    nights = row['Nights']

    subject, body = generate_supplier_email(
        supplier_name=supplier_name,
        num_adults=adults,
        num_children=children,
        accommodation_type=accommodation,
        destination=destination,
        check_in_date=check_in,
        check_out_date=check_out,
        num_nights=nights
    )

    cc_emails = ["cc1@example.com", "cc2@example.com"] 

    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = supplier_email
    msg['Cc'] = ", ".join(cc_emails)  
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            all_recipients = [supplier_email] + cc_emails
            server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
            print(f"✅ Email sent to {supplier_email} with CC to {', '.join(cc_emails)}")
    except Exception as e:
        print(f"❌ Failed to send to {supplier_email}: {e}")

