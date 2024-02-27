from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials 
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import datetime 
import os.path
import re
import base64
from bs4 import BeautifulSoup  # You might need to install this package
SCOPES = ["https://www.googleapis.com/auth/calendar", "https://www.googleapis.com/auth/gmail.readonly", "https://www.googleapis.com/auth/tasks"]

def extract_text_from_html(html_content):
    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Extract text from the parsed HTML
    text = soup.get_text(separator='\n', strip=True)
    
    return text
def format_due_date(due_date_str):
    try:
        # Parse the due date string
        due_date = datetime.datetime.strptime(due_date_str, "%b %d at %I:%M%p")
        # Get the current year
        current_year = datetime.datetime.now().year
        # If the year information is missing, set it to the current year
        if due_date.year == 1900:
            due_date = due_date.replace(year=current_year)
        # Format the due date in RFC3339 format
        return due_date.isoformat() + 'Z'
    except ValueError:
        print("Error: Unable to parse due date.")
        return None
def parse_assignments_from_email(email_body):
    assignments = []
    
    # Preprocess HTML to remove new lines and unnecessary whitespace
    soup = BeautifulSoup(email_body, 'html.parser')
    
    # Find all <b> tags containing assignment details
    assignment_tags = soup.find_all('b', string=re.compile(r"Assignment Created"))
    
    # Iterate over <b> tags and extract assignment details
    for tag in assignment_tags:
        assignment_name = tag.text.strip()  # Extract assignment name
        assignment_name = assignment_name.replace("Assignment Created - ", "")  # Remove prefix
        next_element = tag.find_next('p')  # Find the next <p> tag containing the due date
        due_date_str = next_element.text.strip()  # Extract due date
        
        # Format due date
        due_date_str = due_date_str.replace('due: ', '')  # Remove 'due: ' from the string
        
        # Append assignment details to the list of assignments
        assignments.append((assignment_name, due_date_str))
        
        # Print assignment details
        print("Assignment Name:", assignment_name)
        print("Due Date:", due_date_str)
        print("=" * 30)  # Add separator for readability
    
    return assignments
def get_email_body(service, user_id, message_id):
    """Get the email body of a specific message."""
    try:
        message = service.users().messages().get(userId=user_id, id=message_id, format='full').execute()
        payload = message['payload']
        if 'parts' in payload:
            for part in payload['parts']:
                if part['mimeType'] == 'text/html':
                    data = part['body']['data']
                    html_content = base64.urlsafe_b64decode(data.encode('ASCII')).decode('utf-8')
                    return html_content
        else:
            body = payload['body']
            if 'data' in body:
                data = body['data']
                html_content = base64.urlsafe_b64decode(data.encode('ASCII')).decode('utf-8')
                return html_content
        return None
    except HttpError as error:
        print(f'An error occurred: {error}')



def authenticate_google_api(scopes):
    """Authenticate with Google APIs and return the service object."""
    creds = None
    # Check if token.json exists, which stores the user's access and refresh tokens.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)
    # If there are no (valid) credentials, initiate the login flow.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(r'C:\Users\Christopher\Desktop\automatedCalendar\Programming\Python\AutomatedCalendar\credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for future runs.
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
        print("Authenticated and credentials saved.")  # Print statement for authentication
    else:
        print("Credentials loaded successfully.")  # Print statement for loaded credentials
    return creds
def create_google_task(service, title, due):
    """Create a task in Google Tasks."""
    try:
        task = {
            'title': title,
            'due': due
        }
        
        # Insert the task
        created_task = service.tasks().insert(tasklist='@default', body=task).execute()
        print('Task created successfully:', created_task.get('title'))
        
    except HttpError as error:
        error_details = error.details() if hasattr(error, 'details') else str(error)
        print(f'An error occurred while creating task: {error_details}')
def main():
    try:
        creds = authenticate_google_api(SCOPES)
        gmail_service = build('gmail', 'v1', credentials=creds)
        tasks_service = build('tasks', 'v1', credentials=creds)

        results = gmail_service.users().messages().list(userId='me', q='subject:"Recent Canvas Notifications" is:unread', maxResults=10).execute()
        messages = results.get('messages', [])
        
        print(f"Fetched {len(messages)} messages.")

        if not messages:
            print("No messages found.")
            return

        for message in messages:
            email_body_html = get_email_body(gmail_service, 'me', message['id'])
            if email_body_html:
                print(f"Found email body for message ID: {message['id']}")
                assignments = parse_assignments_from_email(email_body_html)
                print(f"Extracted {len(assignments)} assignments.")
                for assignment_name, due_date_formatted in assignments:
                    print(f"Creating task: {assignment_name} due on {due_date_formatted}")
                    
                    # Format the due date before creating the task
                    formatted_due_date = format_due_date(due_date_formatted)
                    if formatted_due_date:
                        try:
                            create_google_task(tasks_service, assignment_name, due=formatted_due_date)
                        except HttpError as e:
                            print(f"An error occurred while creating task: {e}")
            else:
                print(f"Could not find the email body for message ID: {message['id']}")
    except Exception as ex:
        print(f"An exception occurred: {ex}")

if __name__ == '__main__':
    main()
