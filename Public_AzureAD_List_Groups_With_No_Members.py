import msal
import requests
from colorama import Fore, Style, init
import time
import sys
import threading
import csv
import keyboard
import os

# Initialize colorama
init(autoreset=True)

# Replace these values with your Azure AD app registration details
CLIENT_ID = '____________'
CLIENT_SECRET = '____________'
TENANT_ID = '____________'

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
else:
    print("Error obtaining access token")
    exit()

# Define the Graph API endpoint
graph_api_endpoint = 'https://graph.microsoft.com/v1.0'


# Function to get all AAD groups
def get_groups():
    groups = []
    endpoint = f'{graph_api_endpoint}/groups'
    headers = {
        'Authorization': f'Bearer {token}'
    }
    while endpoint:
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            data = response.json()
            groups.extend(data['value'])
            endpoint = data.get('@odata.nextLink')
        else:
            print("Error fetching groups:", response.status_code, response.text)
            break
    return groups


# Function to check if a group has members
def group_has_members(group_id):
    endpoint = f'{graph_api_endpoint}/groups/{group_id}/members'
    headers = {
        'Authorization': f'Bearer {token}'
    }
    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        data = response.json()
        return len(data['value']) > 0, len(data['value'])
    else:
        print(f"Error fetching members for group {group_id}:", response.status_code, response.text)
        return False, 0


# Function to get the group owner
def get_group_owner(group_id):
    endpoint = f'{graph_api_endpoint}/groups/{group_id}/owners'
    headers = {
        'Authorization': f'Bearer {token}'
    }
    response = requests.get(endpoint, headers=headers)
    if response.status_code == 200:
        data = response.json()
        if data['value']:
            return data['value'][0].get('displayName', 'Unknown')
    return 'Unknown'


# Function to determine the type of the group
def get_group_type(group):
    if group.get('securityEnabled'):
        return 'Security'
    elif 'Unified' in group.get('groupTypes', []):
        return 'Office365'
    else:
        return 'DL'


# Function to get additional group details
def get_group_details(group):
    created_date = group.get('createdDateTime', 'Unknown')
    visibility = group.get('visibility', 'Unknown')
    email = group.get('mail', 'Unknown')
    return created_date, visibility, email


# Animation function
stop_animation = False


def loading_animation():
    global stop_animation
    while not stop_animation:
        for char in ['.', '..', '...']:
            sys.stdout.write('\r' + char)
            sys.stdout.flush()
            time.sleep(0.5)
            if stop_animation:
                break
    sys.stdout.write('\r    \r')
    sys.stdout.flush()


# Function to stop the script
def stop_script():
    global stop_animation, interrupted
    stop_animation = True
    interrupted = True


# Set up key listener for 'Q' key
keyboard.add_hotkey('q', stop_script)


# Get all groups and print those with no members one by one
def list_groups_with_no_members():
    global stop_animation, interrupted
    interrupted = False
    print(Fore.YELLOW + "Groups with no members:")
    print(Fore.YELLOW + "Press 'Q' at any time to abort the script and export the data.")

    # Start the animation thread
    stop_animation = False
    animation_thread = threading.Thread(target=loading_animation)
    animation_thread.start()

    groups = get_groups()
    no_member_groups = []
    for group in groups:
        if interrupted:
            break
        has_members, member_count = group_has_members(group['id'])
        if not has_members:
            # Temporarily stop the animation to print the group name
            stop_animation = True
            animation_thread.join()

            sys.stdout.write('\r' + ' ' * 10 + '\r')  # Clear the line
            group_type = get_group_type(group)
            owner = get_group_owner(group['id'])
            description = group.get('description', 'No description')
            created_date, visibility, email = get_group_details(group)
            print(Fore.GREEN + group[
                'displayName'] + " " + Fore.BLUE + f"({group_type}) " + Fore.MAGENTA + f"Owner: {owner} " + Fore.CYAN + f"Description: {description} " + Fore.YELLOW + f"Created Date: {created_date} " + Fore.LIGHTBLUE_EX + f"Visibility: {visibility} " + Fore.LIGHTGREEN_EX + f"Email: {email} " + Fore.RED + f"Member Count: {member_count}")
            no_member_groups.append(
                (group['displayName'], group_type, owner, description, created_date, visibility, email, member_count))

            # Restart the animation thread
            stop_animation = False
            animation_thread = threading.Thread(target=loading_animation)
            animation_thread.start()

    # Stop the animation thread after processing all groups
    stop_animation = True
    animation_thread.join()

    return no_member_groups


# Function to export the result to CSV
def export_to_csv(group_list):
    downloads_folder = get_downloads_folder()
    file_path = os.path.join(downloads_folder, 'groups_with_no_members.csv')
    with open(file_path, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(
            ['Group Name', 'Group Type', 'Owner', 'Description', 'Created Date', 'Visibility', 'Email', 'Member Count'])
        for group_name, group_type, owner, description, created_date, visibility, email, member_count in group_list:
            writer.writerow([group_name, group_type, owner, description, created_date, visibility, email, member_count])
    print(Fore.CYAN + f"Groups with no members have been exported to your Downloads folder: {file_path}")


# Determine the Downloads folder
def get_downloads_folder():
    if os.name == 'nt':  # Windows
        return os.path.join(os.environ['USERPROFILE'], 'Downloads')
    else:  # macOS and Linux
        return os.path.join(os.path.expanduser('~'), 'Downloads')


# Run the function and prompt for CSV export
try:
    no_member_groups = list_groups_with_no_members()
except KeyboardInterrupt:
    print("\nProcess interrupted by user.")

export_choice = input(Fore.YELLOW + "Do you want to export the result to a CSV file? (yes/no): ").strip().lower()

if export_choice == 'yes':
    export_to_csv(no_member_groups)
else:
    print(Fore.CYAN + "Export canceled.")