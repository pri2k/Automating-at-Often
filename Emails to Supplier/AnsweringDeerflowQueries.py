import os
import time
import requests
import json
import pandas as pd
from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import pickle
from google_auth_oauthlib.flow import InstalledAppFlow

# ─────────────────────────────────────────────────────────────────────────────
# Load Environment Variables and Google Sheets Setup
# ─────────────────────────────────────────────────────────────────────────────

load_dotenv()

SPREADSHEET_ID = os.getenv("SHEET_ID")
RANGE_NAME = "CustomerEnquiry!A1:R1000"

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]


CREDENTIALS_FILE = 'credentials.json'
TOKEN_PICKLE = 'token.pickle'

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

creds = get_user_credentials()
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()

# ─────────────────────────────────────────────────────────────────────────────
# Sync Function: Check, Query Gemini, Write Back
# ─────────────────────────────────────────────────────────────────────────────
def fetch_sheet_data():
    """Fetch customer data from the Google Sheet and return it as a DataFrame."""
    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME
    ).execute()

    values = result.get("values", [])
    filtered_values = [row for row in values if any(cell.strip() for cell in row)]

    if not filtered_values:
        print("⚠️ No data found in the customer sheet.")
        return pd.DataFrame(), RANGE_NAME

    headers = filtered_values[0]  # First row as header
    data_rows = filtered_values[1:]  # Remaining rows

    num_columns = len(headers)

    normalized_rows = [
        row[:num_columns] + [''] * (num_columns - len(row))
        if len(row) != num_columns else row
        for row in data_rows
    ]

    df = pd.DataFrame(normalized_rows, columns=headers)

    values = result.get('values', [])

    return df


def sync_sheet():
    try:
        df = fetch_sheet_data()

        # print(df)
        exit

        for index, row in df.iterrows():
            query = row.get("DeerFlow", "").strip().lower() 
            answer = row.get("Answer", "").strip().lower() 
            if query == "" or answer != "":
                continue  

            print(f"Processing: {query}")

            payload = {
                "messages": [
                    {
                        "role": "user",
                        "content": query
                    }
                ],
                "thread_id": "_default_",
                "max_plan_iterations": 3,        # Increase planning steps for deeper research
                "max_step_num": 7,               # Allow more steps in the AI’s workflow for details
                "auto_accepted_plan": True,
                "interrupt_feedback": "",
                "mcp_settings": {
                    "role": "Concierge and Research Assistant specialized in restaurants and travel",
                    "contextual_goals": [
                        "Find and list restaurants or beach clubs based on cuisine, rating, and location",
                        "Provide accurate and detailed information about reservation processes, pricing, ambiance, and user reviews",
                        "Use latest and trusted travel and restaurant data",
                        "Be polite, clear, and provide thorough explanations with examples or highlights"
                    ],
                    "preferred_sources": ["Google Maps", "Michelin Guide", "TripAdvisor", "Zomato", "Official Restaurant Websites"],
                    "tone": "Polite, helpful, and detailed",
                    "tools": ["web_search", "reservation_info_api", "review_aggregation_api"],
                    "response_format": "JSON with detailed fields: name, location, type, rating, reservation details, ambiance, price range, and brief summary"
                },
                "enable_background_investigation": True,
                "debug": False
            }

                

            r = requests.post("http://localhost:8000/api/chat/stream", json=payload, stream=True)

            if r.status_code == 200:
                print("----- RAW RESPONSE START -----")
                full_response = ""
                for line in r.iter_lines(decode_unicode=True):
                    if line:
                        print(line)
                        if line.startswith("data: "):
                            json_str = line[6:]  # Strip off "data: "
                            # try:
                            data = json.loads(json_str)
                            chunk = data.get("content", "")
                            full_response += chunk
                            # except json.JSONDecodeError:
                            #     print("⚠️ Skipped invalid JSON data line:", line)
                        # else:
                            # print("⚠️ Skipped non-data line:", line)
                print("----- RAW RESPONSE END -----")

                answer = full_response.strip()
                print(f"✅ Final Answer: {answer}")

                headers = df.columns.tolist()
                answer_col_index = headers.index("Answer")  
                column_letter = chr(65 + answer_col_index)    

                update_range = f"CustomerEnquiry!{column_letter}{index+2}"
                body = {
                    "values": [[answer]]
                }
                service.spreadsheets().values().update(
                    spreadsheetId=SPREADSHEET_ID,
                    range=update_range,
                    valueInputOption="RAW",
                    body=body
                ).execute()
                print(f"📝 Wrote answer to {update_range}")

            else:
                print(f"❌ Error from stream API: {r.status_code}")
                print(f"🔄 Response Text: {r.text}")

    except Exception as e:
        print(f"Exception during sync: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# Run Every 10 Seconds
# ─────────────────────────────────────────────────────────────────────────────

print("Start")
while True:
    sync_sheet()
    time.sleep(10)
