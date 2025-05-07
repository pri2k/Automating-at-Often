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
SHEET_RANGE = f'{SHEET_NAME}!A2:L1000'
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
    num_columns = max(len(row) for row in filtered_values)
    normalized_values = [row + [''] * (num_columns - len(row)) for row in filtered_values]

    df = pd.DataFrame(normalized_values, columns=[
        "Customer Name", "Country", "Destination", "Travel Dates", "Number of People",
        "Accommodation Type", "Activities", "Query", "Sent to Supplier", "Supplier Email",
        "Supplier Name", "Supplier Response"
    ])
    return df, SHEET_RANGE

def fetch_supplier_data():
    """Fetch supplier data from the Google Sheet."""
    supplier_range = 'Supplier!B2:E'
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=GOOGLE_SHEET_ID,
        range=supplier_range
    ).execute()

    values = result.get('values', [])
    if not values:
        return pd.DataFrame()

    return pd.DataFrame(values, columns=["Supplier Name", "Email", "Country", "Destination"])

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

def mark_email_sent(sheet, row_index: int):
    """Update the Google Sheet to mark the email as sent."""
    update_range = f"{SHEET_NAME}!{STATUS_COLUMN}{row_index + 2}"  # +2 for header and 1-indexing
    sheet.values().update(
        spreadsheetId=GOOGLE_SHEET_ID,
        range=update_range,
        valueInputOption="USER_ENTERED",
        body={"values": [["Email Sent"]]}
    ).execute()

def process_new_entries():
    """Process the customer requests and send emails to matching suppliers."""
    df, sheet_range = fetch_sheet_data(sheets_service)
    suppliers_df = fetch_supplier_data()

    if df.empty:
        print("No data found in sheet.")
        return

    for index, row in df.iterrows():
        # Skip rows where email is already sent
        if row.get("Sent to Supplier", "").strip().lower() == "email sent":
            continue

        customer_country = row.get("Country", "")
        customer_destination = row.get("Destination", "")
        customer_query = row.get("Query", "")

        # Find a matching supplier based on country and destination
        matching_supplier = suppliers_df[
            (suppliers_df["Country"].str.contains(customer_country, case=False, na=False)) &
            (suppliers_df["Destination"].str.contains(customer_destination, case=False, na=False))
        ]

        if matching_supplier.empty:
            print(f"❌ No matching supplier found for {customer_destination}, {customer_country}")
            continue

        supplier_email = matching_supplier.iloc[0]["Email"]
        supplier_name = matching_supplier.iloc[0]["Supplier Name"]

        # Ensure necessary data is available before sending an email
        if not supplier_email or not customer_query:
            print(f"⚠️ Missing data in row {index + 2}. Skipping.")
            continue

        subject = "Quotation Request for Upcoming Travel Booking"
        email_body = generate_email_for_supplier(supplier_name, customer_query)

        try:
            send_gmail(supplier_email, subject, email_body)
            mark_email_sent(sheets_service, index)
            print(f"✅ Email sent to {supplier_email}")
        except Exception as e:
            print(f"❌ Error sending email to {supplier_email}: {e}")

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
