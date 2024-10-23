import os.path
import base64
import json
import re
import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.text import MIMEText

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.send']

CUMULATIVE_DATA_FILE = 'cumulative_sales.json'

def authenticate_gmail():
    """Authenticate and return the Gmail API service."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)
    return service

def extract_sales_info(email_body):
    """Extract relevant sales information from the email body."""
    sales_info = {}

    # Regular expressions to capture relevant details from the email
    exchange_match = re.search(r"Exchange:\s*(.*)", email_body)
    team_match = re.search(r"Team/Performer:\s*(.*)", email_body)
    venue_match = re.search(r"Venue:\s*(.*)", email_body)
    cost_match = re.search(r"Cost:\s*\$(.*)", email_body)
    payout_match = re.search(r"Payout:\s*\$(.*)", email_body)
    profit_match = re.search(r"Profit:\s*\$(.*)", email_body)

    # Sanitize and convert matched strings to floats by removing commas and stripping whitespace
    sales_info['Exchange'] = exchange_match.group(1) if exchange_match else 'N/A'
    sales_info['Team/Performer'] = team_match.group(1) if team_match else 'N/A'
    sales_info['Venue'] = venue_match.group(1) if venue_match else 'N/A'
    
    sales_info['Cost'] = float(cost_match.group(1).replace(',', '').strip()) if cost_match else 0.0
    sales_info['Payout'] = float(payout_match.group(1).replace(',', '').strip()) if payout_match else 0.0
    sales_info['Profit'] = float(profit_match.group(1).replace(',', '').strip()) if profit_match else 0.0

    return sales_info

def get_sales_emails(service, since_first_of_month=False):
    """Fetches emails from the user's inbox with [TV Sales] in the subject."""
    if since_first_of_month:
        # Get emails from the 1st of the current month
        first_of_month = datetime.date.today().replace(day=1).strftime('%Y/%m/%d')
        query = f'subject:"[TV Sales]" after:{first_of_month}'
    else:
        # Get emails from yesterday and today to account for timezone differences
        yesterday = (datetime.date.today() - datetime.timedelta(days=1)).strftime('%Y/%m/%d')
        query = f'subject:"[TV Sales]" after:{yesterday}'
    
    result = service.users().messages().list(userId='me', q=query).execute()
    messages = result.get('messages', [])
    
    sales_data = []
    if not messages:
        print("No [TV Sales] emails found for the query.")
        return sales_data

    # Get today's date in local time to filter out yesterday's sales for daily totals
    today_start = datetime.datetime.combine(datetime.date.today(), datetime.time(0, 0))

    for msg in messages:
        msg_data = service.users().messages().get(userId='me', id=msg['id']).execute()
        payload = msg_data['payload']
        internal_date = int(msg_data['internalDate']) / 1000  # Convert from milliseconds
        email_date = datetime.datetime.fromtimestamp(internal_date)  # Convert to local datetime

        # If we're fetching cumulative data, include everything from the month
        # If fetching daily data, only include emails from today (after midnight)
        if since_first_of_month or email_date >= today_start:
            if 'parts' in payload:
                for part in payload['parts']:
                    if part['mimeType'] == 'text/plain' and 'data' in part['body']:
                        email_body = base64.urlsafe_b64decode(part['body']['data']).decode('utf-8')
                        sales_info = extract_sales_info(email_body)
                        sales_data.append(sales_info)
            elif 'body' in payload and 'data' in payload['body']:
                email_body = base64.urlsafe_b64decode(payload['body']['data']).decode('utf-8')
                sales_info = extract_sales_info(email_body)
                sales_data.append(sales_info)

    return sales_data

def send_summary_email(service, sales_data_today, cumulative_sales_data):
    """Sends an email summary of the sales for the day."""
    daily_cost = daily_payout = daily_profit = 0.0

    # Start the HTML table with Apple-inspired CSS styling
    summary = """
    <html>
    <head>
    <style>
        table {
            width: 100%;
            border-collapse: collapse;
            border: 1px solid #dddddd;
            font-family: Arial, sans-serif;
        }
        th, td {
            text-align: left;
            padding: 8px;
        }
        th {
            background-color: #f8f8f8;
            font-weight: bold;
            color: #333;
            border-bottom: 1px solid #dddddd;
        }
        td {
            border-bottom: 1px solid #dddddd;
            font-size: 14px;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        h2 {
            font-family: Arial, sans-serif;
            font-weight: bold;
            color: #333;
            border-bottom: 2px solid #eeeeee;
            padding-bottom: 10px;
        }
        .totals {
            font-weight: bold;
            text-align: right;
        }
    </style>
    </head>
    <body>
    <h2>Daily Sales Report</h2>
    <table>
        <tr>
            <th>Sale #</th>
            <th>Exchange</th>
            <th>Team/Performer</th>
            <th>Venue</th>
            <th>Cost</th>
            <th>Payout</th>
            <th>Profit</th>
        </tr>
    """

    # Add rows for each sale today
    for idx, sale in enumerate(sales_data_today, 1):
        profit_color = "green" if sale['Profit'] >= 0 else "red"
        summary += (f"""
        <tr>
            <td>{idx}</td>
            <td>{sale['Exchange']}</td>
            <td>{sale['Team/Performer']}</td>
            <td>{sale['Venue']}</td>
            <td>${sale['Cost']:.2f}</td>
            <td>${sale['Payout']:.2f}</td>
            <td style="color:{profit_color};">${sale['Profit']:.2f}</td>
        </tr>
        """)

        daily_cost += sale['Cost']
        daily_payout += sale['Payout']
        daily_profit += sale['Profit']

    # Add a row for the daily totals
    summary += f"""
    <tr class="totals">
        <td colspan="4" align="right">Today's Totals:</td>
        <td>${daily_cost:.2f}</td>
        <td>${daily_payout:.2f}</td>
        <td style="color:{'green' if daily_profit >= 0 else 'red'}">${daily_profit:.2f}</td>
    </tr>
    """

    # Calculate cumulative totals
    cumulative_cost = cumulative_payout = cumulative_profit = 0.0
    for sale in cumulative_sales_data:
        cumulative_cost += sale['Cost']
        cumulative_payout += sale['Payout']
        cumulative_profit += sale['Profit']

    # Add a row for the cumulative totals
    summary += f"""
    <tr class="totals">
        <td colspan="4" align="right">Cumulative Totals (Month-to-Date):</td>
        <td>${cumulative_cost:.2f}</td>
        <td>${cumulative_payout:.2f}</td>
        <td style="color:{'green' if cumulative_profit >= 0 else 'red'}">${cumulative_profit:.2f}</td>
    </tr>
    """

    # End the HTML table
    summary += """
    </table>
    </body>
    </html>
    """

    # Format the email content as HTML
    message = MIMEText(summary, "html")
    message['to'] = 'skyseats2@gmail.com'  # Updated recipient email address
    message['from'] = 'me'
    message['subject'] = 'Daily Sales Summary'

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
    message = {'raw': raw}
    
    service.users().messages().send(userId='me', body=message).execute()
    print("Summary email sent to skyseats2@gmail.com.")

def main():
    service = authenticate_gmail()

    # Fetch sales emails starting from today for daily totals
    sales_data_today = get_sales_emails(service, since_first_of_month=False)

    # Fetch cumulative sales emails starting from the 1st of the current month
    cumulative_sales_data = get_sales_emails(service, since_first_of_month=True)

    send_summary_email(service, sales_data_today, cumulative_sales_data)

if __name__ == '__main__':
    main()
