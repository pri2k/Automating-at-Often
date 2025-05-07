import os
import time
import base64
from email.mime.text import MIMEText
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]

class SheetProcessor:
    def __init__(self, creds, sheet_url):
        self.creds = creds
        self.sheet_id = sheet_url.split("/d/")[1].split("/")[0]
        self.sheet_service = build('sheets', 'v4', credentials=creds)
        self.gmail_service = build('gmail', 'v1', credentials=creds)

    def process_bookings(self):
        try:
            range_name = 'Sheet1!A1:Z'
            sheet = self.sheet_service.spreadsheets()
            result = sheet.values().get(spreadsheetId=self.sheet_id, range=range_name).execute()
            values = result.get('values', [])

            if not values:
                print("No data found.")
                return

            headers = values[0]
            bookings = [dict(zip(headers, row)) for row in values[1:]]
            updated_rows = []

            for i, booking in enumerate(bookings):
                if booking.get('Email Sent'):
                    continue

                email = booking.get('Supplier Email')
                name = booking.get('Name')
                destination = booking.get('Destination')
                dates = booking.get('Travel Dates')
                pax = booking.get('PAX')
                hotel_pref = booking.get('Hotel Preference')

                email_body = f"""
                Dear Supplier,

                We have a new travel booking request:

                - Client Name: {name}
                - Destination: {destination}
                - Travel Dates: {dates}
                - Number of People (PAX): {pax}
                - Hotel Preference: {hotel_pref}

                Please send your best quotation at your earliest convenience.

                Regards,
                Brighter Beyond Travel Team
                """

                subject = f"Quotation Request: {destination} Trip for {pax} pax"
                self.send_email(email, subject, email_body)

                # Mark "Email Sent" as Yes
                row_index = i + 2  # account for header
                update_range = f"Sheet1!Z{row_index}"  # Assuming column Z is "Email Sent"
                sheet.values().update(
                    spreadsheetId=self.sheet_id,
                    range=update_range,
                    valueInputOption="RAW",
                    body={"values": [["Yes"]]}
                ).execute()
                print(f"✅ Email sent and sheet updated for row {row_index}")

        except HttpError as error:
            print(f"An error occurred: {error}")
    
    def send_email(self, to, subject, body_text):
        message = MIMEText(body_text)
        message['to'] = to
        message['subject'] = subject
        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
        try:
            self.gmail_service.users().messages().send(userId="me", body={"raw": raw}).execute()
            print(f"📧 Email sent to {to}")
        except HttpError as error:
            print(f"❌ Failed to send email to {to}: {error}")

def main():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    sheet_url = 'https://docs.google.com/spreadsheets/d/1bT2I_l6hZwAGKk9Y5JROUnCOpf9fD7WmEY6EnGHwAk8/edit#gid=0'
    processor = SheetProcessor(creds, sheet_url)

    print("📡 Starting continuous scan for new bookings...")
    try:
        while True:
            processor.process_bookings()
            print("✅ Scan complete. Waiting for 60 seconds before next check...\n")
            time.sleep(60)
    except KeyboardInterrupt:
        print("🔴 Script stopped manually.")

if __name__ == '__main__':
    main()
