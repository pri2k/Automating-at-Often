import os
import time
import gspread
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from oauth2client.service_account import ServiceAccountCredentials
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from pathlib import Path
import base64

# Configuration - Load from environment variables or config file
CONFIG = {
    "GOOGLE_SHEETS": {
        "SCOPE": ["https://www.googleapis.com/auth/spreadsheets",
                 "https://www.googleapis.com/auth/drive"],
        "CREDS_FILE": "google_sheets_creds.json",
        "SPREADSHEET_NAME": "Trip Bookings",
        "WORKSHEET_NAME": "Sheet1"
    },
    "GMAIL": {
        "SCOPES": ["https://www.googleapis.com/auth/gmail.send"],
        "TOKEN_FILE": "gmail_token.json",
        "CREDS_FILE": "gmail_creds.json",
        # "SENDER_EMAIL": "your_email@often.club"
        "SENDER_EMAIL": "priya.project.trial@gmail.com"
    },
    "APP": {
        "COMPANY_NAME": "often.club",
        "TEAM_DESCRIPTION": "In addition to a travel platform, we run a concierge desk for select members, dedicated to helping them curate, book, and manage exceptional experiences, including premium hotel stays, trip packages, and events.",
        "DELAY_BETWEEN_EMAILS": 1,
        "DEFAULT_ATTACHMENTS": []
        # "DEFAULT_ATTACHMENTS": ["often_club_policies.pdf"]
    }
}

class EmailService:
    def __init__(self):
        self.creds = self._authenticate_gmail()
        self.service = build('gmail', 'v1', credentials=self.creds)

    def _authenticate_gmail(self):
        """Authenticate with Gmail API using OAuth"""
        creds = None
        if os.path.exists(CONFIG["GMAIL"]["TOKEN_FILE"]):
            creds = Credentials.from_authorized_user_file(
                CONFIG["GMAIL"]["TOKEN_FILE"], CONFIG["GMAIL"]["SCOPES"])
        
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    CONFIG["GMAIL"]["CREDS_FILE"], CONFIG["GMAIL"]["SCOPES"])
                creds = flow.run_local_server(port=0)
            
            with open(CONFIG["GMAIL"]["TOKEN_FILE"], 'w') as token:
                token.write(creds.to_json())
        
        return creds

    def send_email(self, to_email, subject, body, cc_emails=None, attachments=None):
        # """Send email with optional attachments using Gmail API"""
        if cc_emails is None:
            cc_emails = []
        if attachments is None:
            attachments = []

        message = MIMEMultipart()
        message['to'] = to_email
        message['from'] = CONFIG["GMAIL"]["SENDER_EMAIL"]
        message['subject'] = subject
        
        if cc_emails:
            message['cc'] = ', '.join(cc_emails)

        message.attach(MIMEText(body))

        for attachment in attachments:
            with open(attachment, 'rb') as file:
                part = MIMEApplication(
                    file.read(),
                    Name=os.path.basename(attachment)
                )
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment)}"'
            message.attach(part)

        raw_message = base64.urlsafe_b64encode(message.as_bytes())
        raw_message = raw_message.decode()

        try:
            self.service.users().messages().send(
                userId='me',
                body={'raw': raw_message}
            ).execute()
            print(f"✅ Email sent to {to_email}")
            return True
        except Exception as e:
            print(f"❌ Failed to send email: {str(e)}")
            return False

class TripBookingProcessor:
    def __init__(self):
        self.email_service = EmailService()
        self.sheet = self._initialize_google_sheet()

    def _initialize_google_sheet(self):
        """Initialize Google Sheets connection"""
        try:
            creds = ServiceAccountCredentials.from_json_keyfile_name(
                CONFIG["GOOGLE_SHEETS"]["CREDS_FILE"], 
                CONFIG["GOOGLE_SHEETS"]["SCOPE"])
            client = gspread.authorize(creds)
            return client.open(CONFIG["GOOGLE_SHEETS"]["SPREADSHEET_NAME"]).worksheet(CONFIG["GOOGLE_SHEETS"]["WORKSHEET_NAME"])
        except Exception as e:
            print(f"❌ Error initializing Google Sheet: {e}")
            raise

    def generate_email_body(self, booking_data, query_type=None):
        """Generate customized email body based on booking data and query type"""
        # Basic trip information
        adults = booking_data.get('Adults', 0)
        children = booking_data.get('Children', 0)
        children_ages = booking_data.get('Children Ages', '')
        destination = booking_data.get('Destination', '')
        check_in = booking_data.get('Checkin', '')
        check_out = booking_data.get('Checkout', '')
        nights = booking_data.get('Nights', 0)
        rooms = booking_data.get('Room Configuration', '')

        # Start building email
        body = f"""Hello,

I'm part of the team at {CONFIG["APP"]["COMPANY_NAME"]}. {CONFIG["APP"]["TEAM_DESCRIPTION"]}

"""

        # Different email templates based on query type
        if query_type == "room_specific":
            room_type = booking_data.get('Room Type', '')
            body += f"""One of our Luxe-tier members has an upcoming trip to {destination} for which we have shortlisted your property. 
We had a question about the room {room_type}:
{booking_data.get('Question', '')}

Awaiting your reply."""
        
        elif query_type == "family_trip":
            body += f"""We are assisting a client planning a family trip to {destination} and would like to get some additional information to advise them appropriately.

Client group:
{adults} adults and {children} children ({children_ages}) in {rooms}

Questions:
{booking_data.get('Questions', '')}

We have recently had a client stay at the resort and would love to close this booking as well."""
        
        elif query_type == "group_quotation":
            body += f"""A client of ours is planning a trip to {destination}. They are a group of {adults} adults and {children} children ({children_ages}) and are planning to travel from {check_in} to {check_out} ({nights} nights).

We are checking these options for them:
{booking_data.get('Options', '')}

Could you please help us with agent quotations for these options?"""
        
        else:  # Default template
            body += f"""We are planning a trip for a group of {adults} adults{f" and {children} children ({children_ages})" if children else ""} to {destination}.

Trip Details:
• Check-in: {check_in}
• Check-out: {check_out}
• Nights: {nights}
• Room Configuration: {rooms}

{booking_data.get('Additional Notes', '')}"""

        body += f"""

Best regards,
{CONFIG["APP"]["COMPANY_NAME"]} Team"""

        return body

    def process_bookings(self):
        """Process all bookings in the sheet"""
        try:
            bookings = self.sheet.get_all_records()
            
            for idx, booking in enumerate(bookings, start=2):
                try:
                    print(f"\nProcessing booking #{idx-1} for {booking.get('Destination', '')}")
                    
                    # Determine email type based on booking data
                    query_type = self._determine_query_type(booking)
                    subject = self._generate_subject(booking, query_type)
                    body = self.generate_email_body(booking, query_type)
                    
                    # Send email
                    attachments = CONFIG["APP"]["DEFAULT_ATTACHMENTS"]
                    if 'Attachments' in booking and booking['Attachments']:
                        attachments.extend(booking['Attachments'].split(','))
                    
                    self.email_service.send_email(
                        # to_email=booking['Supplier Email'],
                        to_email = "priya.keshrinis@outlook.com",
                        subject=subject,
                        body=body,
                        cc_emails=booking.get('CC Emails', '').split(','),
                        attachments=attachments
                    )
                    
                    time.sleep(CONFIG["APP"]["DELAY_BETWEEN_EMAILS"])
                
                except Exception as e:
                    print(f"❌ Error processing booking #{idx-1}: {str(e)}")
                    continue
        
        except Exception as e:
            print(f"❌ Fatal error in processing: {str(e)}")

    def _determine_query_type(self, booking):
        """Determine the type of query based on booking data"""
        if 'Question' in booking and booking['Question']:
            return "room_specific"
        elif 'Questions' in booking and booking['Questions']:
            return "family_trip"
        elif 'Options' in booking and booking['Options']:
            return "group_quotation"
        return None

    def _generate_subject(self, booking, query_type):
        """Generate appropriate email subject"""
        destination = booking.get('Destination', '')
        adults = booking.get('Adults', 0)
        children = booking.get('Children', 0)
        
        if query_type == "room_specific":
            return f"Question about {booking.get('Room Type', '')} at {destination}"
        elif query_type == "family_trip":
            return f"Family Trip Inquiry for {destination} - {adults+children} PAX"
        elif query_type == "group_quotation":
            return f"Quotation Request for Group Stay at {destination}"
        return f"Trip Inquiry for {destination} - {adults+children} PAX"

if __name__ == "__main__":
    processor = TripBookingProcessor()
    processor.process_bookings()