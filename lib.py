import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd

import yagmail

# For reading sensitive info
import os
from dotenv import load_dotenv

import datetime

load_dotenv()

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = os.getenv("SAMPLE_SPREADSHEET_ID")
SAMPLE_RANGE_NAME = os.getenv("SAMPLE_RANGE_NAME")

def get_data():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("google_sheets_creds/token.json"):
        creds = Credentials.from_authorized_user_file("google_sheets_creds/token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("google_sheets_creds/credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("google_sheets_creds/token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME)
            .execute()
        )

        values = result.get("values", [])

        if not values:
            print("No data found.")
            return

        # At this point, values is a list of lists.
        # First list the header row. Pop this off, and convert to Pandas DF.
        headers = values.pop(0)
        df = pd.DataFrame(values, columns=headers)

        return Sheet(df)

    except HttpError as err:
        print(err)
        return


# Expects a list of strings. Each gets added to the body of the email_template.html, then returned as str
def format_email_html(lst):

    entries = ""
    for entry in lst:
        entries += entry

    with open("email_template.html", "r") as email:
        file_str = ""
        for line in email.readlines():
            file_str += line
        return file_str.replace("<!-- ADD BODY ENTRIES HERE -->", entries)
    
    return None

def send_email(data):
    yag = yagmail.SMTP(os.getenv("HOST_EMAIL"), oauth2_file="./gmail_creds/oauth2_creds.json")
    subject = "Muffim House Reminder Email - "
    subject += datetime.datetime.today().date().strftime("%d %b %Y")
    if data is None:
        yag.send(to=os.getenv("TARGET_EMAIL"), subject=subject, contents="Error reading data from sheet")
    yag.send(to=os.getenv("TARGET_EMAIL"), subject=subject, contents=data)
    yag.close()


class Sheet:
    
    ## Holds data as a pandas dataframe constructed based on structure of the Google Sheet
    def __init__(self, df):
        self.df = df
        # Convert all dates from '01 Jan 1970' format to datetime '1970-01-01' format
        self.df['Due'] = self.df['Due'].apply(lambda x:
            datetime.datetime.strptime(str(x), "%d %b %Y").date())
        self.df['Done'] = self.df['Done'].apply(lambda x:
            datetime.datetime.strptime(str(x), "%d %b %Y").date())

    # Return a printable dataframe of items due between today and @days number of days from now
    def get_due_soon(self, days):
        today = datetime.datetime.today().date()
        later_date = today + datetime.timedelta(days=days)
        filtered_df = self.df.loc[(self.df["Due"] >= today) & (self.df["Due"] < later_date)]
        return self.format(filtered_df)
    
    # Return a printable dataframe of items past due
    def get_overdue(self):
        today = datetime.datetime.today().date()
        filtered_df = self.df.loc[self.df["Due"] < today]
        return self.format(filtered_df)

    def format(self, df):
        filtered_df = df.loc[:,("Activity", "Due")]
        # Convert all dates back from datetime '1970-01-01' format to readable '01 Jan 1970'
        filtered_df['Due'] = filtered_df['Due'].apply(lambda x:
            datetime.datetime.strftime(x, "%d %b %Y"))
        return filtered_df.to_html(index=False)
