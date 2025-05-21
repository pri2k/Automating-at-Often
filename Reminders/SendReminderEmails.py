from __future__ import print_function
import datetime
import os.path
import base64
from email.mime.text import MIMEText

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]
SPREADSHEET_ID = os.getenv("SHEET_ID")  
SHEET_NAME = "SendReminders"
RANGE_NAME = 'SHEET_NAME!A2:G'  
YOUR_EMAIL = 'priya.project.trial@gmail.com'
REMINDER_DAYS = [60, 30, 7, 6, 5, 4, 3, 2, 1]

# ─────────────────────────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────────────────────────

def get_credentials():
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

    return creds

# ─────────────────────────────────────────────────────────────────────────────
# Send Email
# ─────────────────────────────────────────────────────────────────────────────

def send_email(service, subject, body, to_email):
    message = MIMEText(body)
    message['to'] = to_email
    message['from'] = YOUR_EMAIL
    message['subject'] = subject

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    body = {'raw': raw}
    service.users().messages().send(userId='me', body=body).execute()

# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    creds = get_credentials()

    # Google Sheets & Gmail services
    sheets_service = build('sheets', 'v4', credentials=creds)
    gmail_service = build('gmail', 'v1', credentials=creds)

    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
        return

    today = datetime.date.today()

    for i, row in enumerate(values, start=2):
        try:
            name = row[0]
            destination = row[2]
            checkin_str = row[3]
            status = row[5] if len(row) > 5 else ''
            last_reminder = row[6] if len(row) > 6 else ''

            if status.lower() == "bookings done":
                continue

            checkin_date = datetime.datetime.strptime(checkin_str, "%Y-%m-%d").date()
            days_left = (checkin_date - today).days

            if days_left in REMINDER_DAYS:
                subject = f"Reminder: Booking for {name} to {destination} on {checkin_str}"
                body = f"{name} is going to {destination} on {checkin_str}. Please confirm whether you have done their booking or not and update it."

                send_email(gmail_service, subject, body, YOUR_EMAIL)

                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                update_range = f"Sheet1!G{i}"
                sheet.values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="RAW",
                    body={"values": [[timestamp]]}
                ).execute()

                print(f"Reminder sent for {name} (Row {i})")

        except Exception as e:
            print(f"Error processing row {i}: {e}")

if __name__ == '__main__':
    main()
