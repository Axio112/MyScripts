import msal
import requests
from colorama import Fore, Style, init
from tqdm import tqdm
import time
import sys
import threading
import csv
import keyboard
import os

# Initialize colorama
init(autoreset=True)

# Replace these values with your Azure AD app registration details
CLIENT_ID = '________________________'
CLIENT_SECRET = '________________________'
TENANT_ID = '________________________'

# Initialize the MSAL confidential client application
app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=f'https://login.microsoftonline.com/{TENANT_ID}',
    client_credential=CLIENT_SECRET
)

# Get the access token
result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

if "access_token" in result:
    token = result["access_token"]
    print("Access token obtained successfully.")
else:
    print("Error obtaining access token:")
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))
    exit()

# Define the Graph API endpoint
graph_api_endpoint = 'https://graph.microsoft.com/v1.0'


# Function to get all users
def get_users():
    users = []
    endpoint = f'{graph_api_endpoint}/users'
    headers = {
        'Authorization': f'Bearer {token}'
    }
    while endpoint:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            data = response.json()
            users.extend(data['value'])
            endpoint = data.get('@odata.nextLink')
        else:
            print("Error fetching users:", response.status_code, response.text)
            break
    return users


# Function to get user last sign-in date
def get_last_signin_date(user_id):
    endpoint = f'{graph_api_endpoint}/auditLogs/signIns?$filter=userId eq \'{user_id}\'&$top=1&$orderby=createdDateTime desc'
    headers = {
        'Authorization': f'Bearer {token}'
    }
    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        data = response.json()
        if data['value']:
            return data['value'][0].get('createdDateTime', 'Unknown')
    return 'Never Signed In'


# Function to stop the script
def stop_script():
    global interrupted
    interrupted = True


# Set up key listener for 'Q' key
keyboard.add_hotkey('q', stop_script)


# Determine the Downloads folder
def get_downloads_folder():
    if os.name == 'nt':  # Windows
        return os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:  # macOS and Linux
        return os.path.join(os.path.expanduser('~'), 'Downloads')


# Get all users and print their information
def list_users():
    global interrupted
    interrupted = False
    print(Fore.YELLOW + "Getting users from directory:")
    print(Fore.YELLOW + "Press 'Q' at any time to abort the script and save the data to CSV.")

    # Prepare CSV file
    downloads_folder = get_downloads_folder()
    file_path = os.path.join(downloads_folder, 'users_in_directory.csv')
    with open(file_path, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(
            ['Display Name', 'User Principal Name', 'Email', 'Account Enabled', 'Created Date', 'Last Sign-In Date',
             'Department', 'Job Title'])

        users = get_users()
        total_users = len(users)

        with tqdm(total=total_users, desc="Processing users", unit="user") as pbar:
            try:
                for user in users:
                    if interrupted:
                        break

                    display_name = user.get('displayName', 'Unknown')
                    user_principal_name = user.get('userPrincipalName', 'Unknown')
                    email = user.get('mail', 'Unknown')
                    account_enabled = user.get('accountEnabled', 'Unknown')
                    created_date = user.get('createdDateTime', 'Unknown')
                    last_signin_date = get_last_signin_date(user['id'])
                    department = user.get('department', 'Unknown')
                    job_title = user.get('jobTitle', 'Unknown')

                    writer.writerow(
                        [display_name, user_principal_name, email, account_enabled, created_date, last_signin_date,
                         department, job_title])

                    pbar.update(1)
            except KeyboardInterrupt:
                print("\nProcess interrupted by user.")
                interrupted = True

    print(Fore.CYAN + f"User details have been exported to your Downloads folder: {file_path}")


# Run the function
list_users()