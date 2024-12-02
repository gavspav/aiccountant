from __future__ import print_function
import os.path
import base64
import email
import re
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import pandas as pd
from email.utils import parsedate_to_datetime
from datetime import datetime
import pytz
from tqdm import tqdm

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def convert_to_utc(dt):
    """Convert datetime to UTC and ensure it's timezone-aware."""
    if dt.tzinfo is None:
        # If datetime is naive, assume it's UTC
        return pytz.utc.localize(dt)
    return dt.astimezone(pytz.utc)

def main():
    """Shows basic usage of the Gmail API."""
    creds = None
    # The file token.json stores the user's access and refresh tokens
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    print("Searching for emails...")
    # Search for emails with specific keywords in subject between dates
    query = "subject:(order OR invoice OR receipt) after:2023/04/01 before:2024/04/10"
    
    # Get all messages using pagination
    messages = []
    next_page_token = None
    page_count = 0
    
    with tqdm(desc="Fetching message list", unit="page") as pbar:
        while True:
            results = service.users().messages().list(
                userId='me', 
                q=query,
                pageToken=next_page_token,
                maxResults=500  # Maximum allowed per request
            ).execute()
            
            if 'messages' in results:
                messages.extend(results['messages'])
            
            page_count += 1
            pbar.update(1)
            pbar.set_postfix({"messages": len(messages)})
            
            next_page_token = results.get('nextPageToken')
            if not next_page_token:
                break

    if not messages:
        print('No messages found.')
        return

    print(f'\nProcessing {len(messages)} messages...')
    transactions = []
    
    # Process messages with progress bar
    for message in tqdm(messages, desc="Processing emails", unit="email"):
        try:
            msg = service.users().messages().get(userId='me', id=message['id'], format='raw').execute()
            msg_str = base64.urlsafe_b64decode(msg['raw'].encode('ASCII'))
            mime_msg = email.message_from_bytes(msg_str)

            # Extract email headers directly from mime_msg
            subject = mime_msg['Subject']
            date_str = mime_msg['Date']
            from_email = mime_msg['From']

            # Convert email date to datetime
            try:
                date = parsedate_to_datetime(date_str)
                date = convert_to_utc(date)  # Ensure datetime is UTC and timezone-aware
            except Exception:
                tqdm.write(f"Could not parse date: {date_str}")
                continue

            # Extract the sender name and email from the From field
            supplier = from_email
            if '<' in from_email:
                supplier = from_email.split('<')[0].strip()
                if not supplier:  # If name part is empty, use the email
                    supplier = from_email.split('<')[1].rstrip('>')

            # Extract email body
            body = ""
            if mime_msg.is_multipart():
                for part in mime_msg.walk():
                    if part.get_content_type() == 'text/plain':
                        try:
                            payload = part.get_payload(decode=True)
                            charset = part.get_content_charset() or 'iso-8859-1'
                            body = payload.decode(charset, errors='replace')
                            break
                        except Exception as e:
                            tqdm.write(f"Error decoding email body: {e}")
                            continue
            else:
                try:
                    payload = mime_msg.get_payload(decode=True)
                    charset = mime_msg.get_content_charset() or 'iso-8859-1'
                    body = payload.decode(charset, errors='replace')
                except Exception as e:
                    tqdm.write(f"Error decoding email body: {e}")
                    continue

            # Parse the body to extract amount
            amount = extract_amount(body)

            transactions.append({
                'Date': date,
                'Amount': amount,
                'Supplier': supplier,
                'Subject': subject
            })
        except Exception as e:
            tqdm.write(f"Error processing message: {e}")
            continue

    print("\nPreparing final output...")
    # Convert to DataFrame and ensure all dates are in UTC
    df = pd.DataFrame(transactions)
    
    if len(df) == 0:
        print("No transactions were successfully processed.")
        return
        
    # Convert any naive datetimes to UTC
    df['Date'] = df['Date'].apply(convert_to_utc)
    
    # Sort by date
    df = df.sort_values('Date')
    
    # Format the date column for CSV output (ISO format)
    df['Date'] = df['Date'].dt.strftime('%Y-%m-%d %H:%M:%S %z')
    
    # Save to CSV
    df.to_csv('gmail_transactions.csv', index=False)
    print(f'\nSaved {len(df)} transactions to gmail_transactions.csv')
    
    # Print summary
    print("\nSummary:")
    print(f"Total emails found: {len(messages)}")
    print(f"Successfully processed: {len(df)}")
    print(f"Failed/Skipped: {len(messages) - len(df)}")

def extract_amount(body_text):
    # Implement parsing logic to extract the amount
    # Example regex for amount in GBP
    match = re.search(r'£\s*([\d,]+\.\d{2})', body_text)
    if match:
        return float(match.group(1).replace(',', ''))
    else:
        return None

if __name__ == '__main__':
    main()