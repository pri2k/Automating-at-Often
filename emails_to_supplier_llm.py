import os
import time
import base64
import pickle
import pandas as pd
from email.mime.text import MIMEText
from dotenv import load_dotenv
import google.generativeai as genai

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

load_dotenv()

# Load environment variables
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")  # Supplier Sheet
SHEET_ID = os.getenv("SHEET_ID")  # Customer Sheet

# Define API Scopes for Google Sheets and Gmail
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]

# File paths for credentials and token
CREDENTIALS_FILE = 'credentials.json'
TOKEN_PICKLE = 'token.pickle'

# Sheet configuration
SHEET_NAME = "CustomerEnquiry"
SHEET_RANGE = f'{SHEET_NAME}!A1:L1000'
STATUS_COLUMN = 'A'
CHECK_INTERVAL_SECONDS = 30  # Polling interval for new entries

# ─────────────────────────────────────────────────────────────────────────────
# Setup: Google Services & Gemini
# ─────────────────────────────────────────────────────────────────────────────

def get_user_credentials():
    """Retrieve or refresh user credentials for Google Sheets and Gmail APIs."""
    creds = None
    if os.path.exists(TOKEN_PICKLE):
        with open(TOKEN_PICKLE, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PICKLE, 'wb') as token:
            pickle.dump(creds, token)

    return creds

# Initialize credentials and services
creds = get_user_credentials()
sheets_service = build('sheets', 'v4', credentials=creds)
gmail_service = build('gmail', 'v1', credentials=creds)

# Initialize Gemini AI model for email generation
genai.configure(api_key=GEMINI_API_KEY)
gemini_model = genai.GenerativeModel('gemini-2.0-flash')

# ─────────────────────────────────────────────────────────────────────────────
# Helper Functions
# ─────────────────────────────────────────────────────────────────────────────

def column_index_to_letter(index):
    """Convert a column index to an Excel-style column letter."""
    result = ''
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result

# ─────────────────────────────────────────────────────────────────────────────
# Main Functions
# ─────────────────────────────────────────────────────────────────────────────

def fetch_sheet_data(service):
    """Fetch customer data from the Google Sheet and return it as a DataFrame."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SHEET_ID,
        range=SHEET_RANGE
    ).execute()

    values = result.get("values", [])
    filtered_values = [row for row in values if any(cell.strip() for cell in row)]

    if not filtered_values:
        print("⚠️ No data found in the customer sheet.")
        return pd.DataFrame(), SHEET_RANGE

    headers = filtered_values[0]  # First row as header
    data_rows = filtered_values[1:]  # Remaining rows

    num_columns = len(headers)

    normalized_rows = [
        row[:num_columns] + [''] * (num_columns - len(row))
        if len(row) != num_columns else row
        for row in data_rows
    ]

    df = pd.DataFrame(normalized_rows, columns=headers)


    print("\n📄 Customer Sheet Data Preview:")
    print(df.head())  
    print("\n📊 Customer Sheet Columns:")
    print(df.columns.tolist())
    # print(f"📦 sheet_data type: {type(df)}")  # Prints the type of the object
    # print(f"📦 sheet_data: {df}")  # Prints the whole result to check its structure

    values = result.get('values', [])

    return df, SHEET_RANGE



def fetch_supplier_data():
    sheet_id = os.getenv("SHEET_ID")
    sheet_range = "Supplier!A1:E"  
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id, range=sheet_range
    ).execute()

    values = result.get("values", [])

    if not values:
        print("⚠️ No data found in the supplier sheet.")
        return pd.DataFrame(), sheet_range

    headers = values[0]
    data_rows = values[1:]

    num_columns = len(headers)

    normalized_rows = [
        row[:num_columns] + [''] * (num_columns - len(row))
        if len(row) != num_columns else row
        for row in data_rows
    ]

    df = pd.DataFrame(normalized_rows, columns=headers)

    print("\n📄 Supplier Sheet Data Preview:")
    print(df)

    print("\n📊 Supplier Sheet Columns:")
    print(df.columns.tolist())

    return df, sheet_range



def generate_email_for_supplier(supplier_name: str, customer_query: str) -> str:
    """Generate a personalized email to the supplier requesting a quote for the customer."""
    prompt = f"""
    You are an assistant at a travel agency. A customer has requested a trip quote.
    Write a polite and professional email to the supplier requesting details and quotation.

    Supplier Name: {supplier_name}
    Customer Query: {customer_query}

    Structure the email with a greeting, a clear explanation of the request, and a polite closing.
    """
    response = gemini_model.generate_content(prompt)
    return response.text.strip()


def send_gmail(to_email: str, subject: str, body: str):
    """Send an email via Gmail API."""
    message = MIMEText(body)
    message['to'] = to_email
    message['subject'] = subject
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    gmail_service.users().messages().send(
        userId="me",
        body={'raw': raw_message}
    ).execute()


def mark_email_sent(sheets_service, row_index: int):
    """Update the Google Sheet to mark the email as sent."""
    update_range = f"{SHEET_NAME}!A{row_index + 2}"  # 'Sent to Supplier' is in column A
    timestamp_range = f"{SHEET_NAME}!B{row_index + 2}"  # 'Email Sent Timestamp' is in column B

    from datetime import datetime
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Update both status and timestamp
    sheets_service.spreadsheets().values().batchUpdate(  #
        spreadsheetId=GOOGLE_SHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": update_range, "values": [["Email Sent"]]},
                {"range": timestamp_range, "values": [[now]]}
            ]
        }
    ).execute()



def process_new_entries():
    """Process the customer requests and send emails to matching suppliers."""
    df, sheet_range = fetch_sheet_data(sheets_service)
    suppliers_df, _ = fetch_supplier_data()

    if df.empty:
        print("❌ No customer data found.")
        return

    for index, row in df.iterrows():
        if row.get("Sent to Supplier", "").strip().lower() == "email sent":
            continue

        customer_country = row.get("Country", "").strip()
        customer_destination = row.get("Destination", "").strip()
        customer_query = row.get("Query", "").strip()

        if not customer_country or not customer_destination:
            print(f"⚠️ Missing critical data in row {index + 2}. Skipping.")
            continue

        matching_supplier = suppliers_df[
            (suppliers_df["Country"].str.contains(customer_country, case=False, na=False)) &
            (suppliers_df["Destination"].str.contains(customer_destination, case=False, na=False))
        ]

        if matching_supplier.empty:
            print(f"❌ No matching supplier for {customer_destination}, {customer_country} in row {index + 2}")
            continue

        supplier_email = matching_supplier.iloc[0]["Email"]
        supplier_name = matching_supplier.iloc[0]["Supplier Name"]

        if not supplier_email:
            print(f"⚠️ Missing supplier email for row {index + 2}. Skipping.")
            continue

        subject = "Quotation Request for Upcoming Travel Booking"
        email_body = generate_email_for_supplier(supplier_name, customer_query)

        try:
            send_gmail(supplier_email, subject, email_body)
            mark_email_sent(sheets_service, index)
            print(f"✅ Email sent to {supplier_email} for row {index + 2}")
        except Exception as e:
            print(f"❌ Error sending email to {supplier_email} in row {index + 2}: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# Main Execution Loop
# ─────────────────────────────────────────────────────────────────────────────

def main():
    """Main function to monitor Google Sheets and send emails based on new entries."""
    print("📡 Monitoring Google Sheet for new supplier requests...")
    while True:
        process_new_entries()
        time.sleep(CHECK_INTERVAL_SECONDS)

if __name__ == "__main__":
    main()