import os
import time
import base64
import pickle
import pandas as pd
from email.mime.text import MIMEText
from dotenv import load_dotenv
import google.generativeai as genai
from datetime import datetime

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ─────────────────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────────────────

load_dotenv()

# Load environment variables
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
SHEET_ID = os.getenv("SHEET_ID")  

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
SHEET_RANGE = f'{SHEET_NAME}!A1:P1000'
STATUS_COLUMN = 'A'
CHECK_INTERVAL_SECONDS = 10  # Polling interval for new entries

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


    # print("\n📄 Customer Sheet Data Preview:")
    # print(df.head())  
    # print("\n📊 Customer Sheet Columns:")
    # print(df.columns.tolist())
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

    # print("\n📄 Supplier Sheet Data Preview:")
    # print(df)

    # print("\n📊 Supplier Sheet Columns:")
    # print(df.columns.tolist())

    return df, sheet_range



def generate_email_for_supplier(supplier_name: str, customer_data: dict) -> str:
    """Generate an email to the supplier using Gemini, with customer data extracted directly."""
    
    # Extract customer data from the provided dictionary
    country = customer_data.get("Country", "N/A")
    destination = customer_data.get("Destination", "N/A")
    lead_name = customer_data.get("Lead Passenger Name", "N/A")
    adults = customer_data.get("Adults", "0")
    children = customer_data.get("Children", "0")
    accommodation = customer_data.get("Accommodation Type", "N/A")
    checkin = customer_data.get("Checkin", "N/A")
    checkout = customer_data.get("Checkout", "N/A")
    nights = customer_data.get("Nights", "N/A")
    activities = customer_data.get("Activities", "Not specified").strip()
    query = customer_data.get("Query", "No additional query provided").strip()
    wants_quote = customer_data.get("Quote", "").strip().lower() == "yes"

    with open("email_examples.txt", "r") as f:
        content = f.read()


    # Build structured prompt for Gemini
    prompt = f"""
    You are Priya, a Travel Consultant at Often. Write a friendly and professional email to a travel supplier named {supplier_name}. This is for a customer trip inquiry. You must not use Markdown or asterisks for formatting — instead, use HTML tags like <b> for emphasis.


    Here are the key details you must naturally include in the email:

    - Guest Name: {lead_name}
    - Country: {country}
    - Destination: {destination}
    - Check-in Date: {checkin}
    - Accommodation Type: {accommodation}
    - Number of Adults: {adults}
    - Number of Children: {children} (mention only if > 0)
    
    If any activities are there, mention them like this: {activities} if it is N/A DO NOT MENTION
    If there is a query, please highlight it like this and ask those queries in a natural human language: <b>{query}</b> if it is N/A DO NOT MENTION

    Here are some example emails (with formatting):
    {content}

    """    

    if wants_quote:
        prompt += "Ask the supplier for a detailed quotation for this trip, including accommodation options, activities (if any), and a price breakdown.\n"
    else:
        prompt += "Do not ask the supplier for quotation. Do not ask about money or price breakdown."

    prompt += "Now, based on this information, generate an email in HTML format (no Markdown)."

    # Send the structured prompt to Gemini and get the response
    # print("Prompt is")
    # print(prompt)
    response = gemini_model.generate_content(prompt)
    # print("Response is")
    # print(response)

    raw_text = response.text

    # print(raw_text)

    if raw_text.startswith("```html"):
        raw_text = raw_text[len("```html\n"):]
    if raw_text.endswith("```"):
        raw_text = raw_text[:-3]

    clean_email_html = raw_text.strip()

    # print(clean_email_html)

    return clean_email_html



def send_gmail(to_email: str, subject: str, body: str):
    """Send an HTML email via Gmail API."""
    message = MIMEText(body, 'html')
    message['To'] = to_email
    message['From'] = "me"  
    message['Subject'] = subject

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

    gmail_service.users().messages().send(
        userId="me",
        body={'raw': raw_message}
    ).execute()


def mark_email_sent(sheets_service, row_index: int):
    """Update the Google Sheet to mark the email as sent."""
    update_range = f"{SHEET_NAME}!A{row_index + 2}"  
    timestamp_range = f"{SHEET_NAME}!B{row_index + 2}"  

    from datetime import datetime
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Update both status and timestamp
    sheets_service.spreadsheets().values().batchUpdate(  
        spreadsheetId=SHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": update_range, "values": [["Email Sent"]]},
                {"range": timestamp_range, "values": [[now]]}
            ]
        }
    ).execute()

def confirm_checkin_email_sent(sheets_service, row_index: int):
    """Update the Google Sheet to mark the email as sent."""
    update_range = f"{SHEET_NAME}!D{row_index + 2}"  

    # Update both status and timestamp
    sheets_service.spreadsheets().values().batchUpdate(  
        spreadsheetId=SHEET_ID,
        body={
            "valueInputOption": "USER_ENTERED",
            "data": [
                {"range": update_range, "values": [["Email Sent"]]},
            ]
        }
    ).execute()

def generate_email_for_checkin(supplier_name: str, customer_data: dict) -> str:
    """Generate an email to the supplier using Gemini, with customer data extracted directly."""
    
    # Extract customer data from the provided dictionary
    country = customer_data.get("Country", "N/A")
    destination = customer_data.get("Destination", "N/A")
    lead_name = customer_data.get("Lead Passenger Name", "N/A")
    adults = customer_data.get("Adults", "0")
    children = customer_data.get("Children", "0")
    accommodation = customer_data.get("Accommodation Type", "N/A")
    checkin = customer_data.get("Checkin", "N/A")
    checkout = customer_data.get("Checkout", "N/A")
    nights = customer_data.get("Nights", "N/A")
    activities = customer_data.get("Activities", "N/A").strip()
    query = customer_data.get("Query", "N/A").strip()

    print("Query is")
    print(query)
    print("Activities are")
    print(activities)


    # Build structured prompt for Gemini
   # Build structured prompt for Gemini
    prompt = f"""
    You are Priya, a Travel Concierge at Often.club. Write a warm, professional email to a travel supplier named {supplier_name}. The purpose of this email is to coordinate a special check-in experience for a Luxe member staying at their property.

    Do not use Markdown or asterisks for formatting — instead, use HTML tags like <b> and <i> if emphasis is needed. Maintain a polite and courteous tone.

    Here are the key details you must naturally include in the email:

    - Guest Name: {lead_name}
    - Country: {country}
    - Destination: {destination}
    - Check-in Date: {checkin}
    - Accommodation Type: {accommodation}
    - Number of Adults: {adults}
    - Number of Children: {children} (mention only if > 0)
    - Special occasion or customer query: <b>{query}</b> (if N/A DO NOT MENTION)
    - Requested activities or support: {activities} (if N/A DO NOT MENTION)

    The email should:
    - Briefly introduce yourself and the Often.club concierge service
    - Mention that this is a Luxe-tier member
    - Kindly ask the supplier to help enhance the guest's stay by arranging appropriate welcome gestures or setup
    - Express gratitude and willingness to provide more information if needed

    Now, based on this information, generate a professional, human-sounding email in HTML format (not Markdown).
    """


    prompt += "Now, based on this information, generate an email in HTML format (no Markdown)."

    # print("Prompt is")
    # print(prompt)
    response = gemini_model.generate_content(prompt)
    # print("Response is")
    # print(response)

    raw_text = response.text

    # print(raw_text)

    if raw_text.startswith("```html"):
        raw_text = raw_text[len("```html\n"):]
    if raw_text.endswith("```"):
        raw_text = raw_text[:-3]

    clean_email_html = raw_text.strip()

    # print(clean_email_html)

    return clean_email_html



def process_old_queries():
    df, sheet_range = fetch_sheet_data(sheets_service)
    suppliers_df, _ = fetch_supplier_data()
    
    today = datetime.now().date()

    for index, row in df.iterrows():
        status = row.get("Confirm Checkin", "").strip().lower()
        if status in ["email sent"]:
            continue

        scheduled_date_str = row.get("Scheduled Date", "").strip()
        scheduled_date = None

        if scheduled_date_str:
            try:
                scheduled_date = datetime.strptime(scheduled_date_str, "%Y-%m-%d").date()
            except ValueError:
                print(f"⚠️ Invalid scheduled date format in row {index + 2}: {scheduled_date_str}. Skipping.")
                continue
        
        scheduled_date = datetime.strptime(scheduled_date_str, "%Y-%m-%d").date()
        
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

        subject = "Upcoming Guest Check-in: Coordination Request"
        email_body = generate_email_for_checkin(supplier_name, row.to_dict())

        try:
            if scheduled_date and scheduled_date == today:
                send_gmail(supplier_email, subject, email_body)
                confirm_checkin_email_sent(sheets_service, index)
                print(f"☑️ Email sent to {supplier_email} for row {index + 2}")
        except Exception as e:
            print(f"❌ Error sending/scheduling email to {supplier_email} in row {index + 2}: {e}")

def process_new_entries():
    """Process the customer requests and send emails to matching suppliers."""
    df, sheet_range = fetch_sheet_data(sheets_service)
    suppliers_df, _ = fetch_supplier_data()

    if df.empty:
        print("❌ No customer data found.")
        return
    

    for index, row in df.iterrows():
        status = row.get("Sent to Supplier", "").strip().lower()
        if status in ["email sent"]:
            continue

        scheduled_date_str = row.get("Scheduled Date", "").strip()
        scheduled_date = None
        if scheduled_date_str:
            try:
                scheduled_date = datetime.strptime(scheduled_date_str, "%Y-%m-%d").date()
            except ValueError:
                print(f"⚠️ Invalid scheduled date format in row {index + 2}: {scheduled_date_str}. Skipping.")
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

        subject = "Request for Upcoming Travel Booking"
        email_body = generate_email_for_supplier(supplier_name, row.to_dict())

        try:
            send_gmail(supplier_email, subject, email_body)
            mark_email_sent(sheets_service, index)
            print(f"✅ Email sent to {supplier_email} for row {index + 2}")
        except Exception as e:
            print(f"❌ Error sending/scheduling email to {supplier_email} in row {index + 2}: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# Main Execution Loop
# ─────────────────────────────────────────────────────────────────────────────

def main():
    """Main function to monitor Google Sheets and send emails based on new entries."""
    print("📡 Monitoring Google Sheet for new supplier requests...")
    while True:
        process_new_entries()
        time.sleep(10)
        process_old_queries()
        time.sleep(CHECK_INTERVAL_SECONDS)

if __name__ == "__main__":
    main()