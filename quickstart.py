import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import os.path 
import base64 
import re

import requests
import io
import pandas as pd    


# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def get_links():
  """Shows basic usage of the Gmail API.
  Lists the user's Gmail labels.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())
  links = []
  try:
    # Call the Gmail API
    service = build("gmail", "v1", credentials=creds)
    # results = service.users().labels().list(userId="me").execute()
    results = service.users().messages().list(userId='me').execute()
    print(results)
    # labels = results.get("labels", [])
    messages = results.get('messages',[])

    # iterate through all the messages 
    for msg in messages: 
        # Get the message from its id 
        txt = service.users().messages().get(userId='me', id=msg['id']).execute() 
  
        # Use try-except to avoid any Errors 
        try: 
            # Get value of 'payload' from dictionary 'txt' 
            payload = txt['payload'] 
            headers = payload['headers'] 
  
            # Look for Subject and Sender Email in the headers 
            for d in headers: 
                if d['name'] == 'Subject': 
                    subject = d['value'] 
                if d['name'] == 'From': 
                    sender = d['value'] 
  
            # The Body of the message is in Encrypted format. So, we have to decode it. 
            # Get the data and decode it with base 64 decoder. 
            parts = payload.get('parts')[0] 
            data = parts['body']['data'] 
            data = data.replace("-","+").replace("_","/") 
            decoded_data = base64.b64decode(data).decode('utf-8')
            # Looks for the link to the excel file
            x = re.search("<https://space\S*token\S*>", decoded_data) 
            link = x.group(0)[1:-1]
            # Downloads the file and puts it in a DataFrame
            get_df(link)
            links.append(link)
  
        except Exception as e:
            print('error', e) 
            pass      
  except HttpError as error:
    # TODO(developer) - Handle errors from gmail API.
    print(f"An error occurred: {error}")
  return links


#Downloads the excel file and puts it in a DataFrame
def get_df(link):
  s=requests.get(link).content
  df = pd.read_excel(io.BytesIO(s), sheet_name='KPI - Meals Saved', engine="openpyxl")
  print(df)

if __name__ == "__main__":
  get_links()