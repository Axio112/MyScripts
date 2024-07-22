import requests
from datetime import datetime, timedelta
import json
from colorama import Fore, Style, init
from tabulate import tabulate
import csv

# Initialize colorama
init(autoreset=True)

def get_access_token(tenant_id, client_id, client_secret):
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default"
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json()["access_token"]

def get_audit_logs(access_token, user_upn, max_pages=10):
    url = "https://graph.microsoft.com/v1.0/auditLogs/directoryAudits"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    thirty_days_ago = (datetime.now() - timedelta(days=30)).isoformat() + 'Z'

    params = {
        "$filter": f"activityDateTime ge {thirty_days_ago} and activityDisplayName eq 'Remove member from group' and (initiatedBy/user/userPrincipalName eq '{user_upn}' or targetResources/any(t:t/userPrincipalName eq '{user_upn}'))",
        "$top": 999
    }

    all_logs = []
    page_count = 0

    while url and page_count < max_pages:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        if "value" in data:
            all_logs.extend(data["value"])

        url = data.get("@odata.nextLink")
        params = {}
        page_count += 1

    return all_logs

def export_to_csv(logs, filepath):
    with open(filepath, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        writer.writerow(["Actor", "OID", "GDN", "Date"])
        for log in logs:
            writer.writerow(log)

def re_add_user_to_groups(access_token, user_upn, group_ids):
    url_template = "https://graph.microsoft.com/v1.0/groups/{group_id}/members/$ref"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    user_id = get_user_id(access_token, user_upn)
    for group_id in group_ids:
        if not is_user_in_group(access_token, user_id, group_id):
            url = url_template.format(group_id=group_id)
            data = {
                "@odata.id": f"https://graph.microsoft.com/v1.0/users/{user_id}"
            }
            response = requests.post(url, headers=headers, json=data)
            if response.status_code == 204:
                print(f"Successfully added {user_upn} to group {group_id}")
            else:
                print(f"Failed to add {user_upn} to group {group_id}: {response.text}")
        else:
            print(f"{user_upn} is already a member of group {group_id}")

def get_user_id(access_token, user_upn):
    url = f"https://graph.microsoft.com/v1.0/users/{user_upn}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    data = response.json()
    return data['id']

def is_user_in_group(access_token, user_id, group_id):
    url = f"https://graph.microsoft.com/v1.0/groups/{group_id}/members/{user_id}/$ref"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.get(url, headers=headers)
    return response.status_code == 200


def main():
    # Replace these with your actual values
    tenant_id = "tenant_id"
    client_id = "client_id"
    client_secret = "client_secret"

    user_upn = input("Enter the User Principal Name (UPN) to look up: ")

    print("Attempting to get access token...")
    access_token = get_access_token(tenant_id, client_id, client_secret)
    print(f"Access token received: {access_token[:10]}...")

    print("Attempting to get audit logs...")
    logs = get_audit_logs(access_token, user_upn, max_pages=10)

    if not logs:
        print(f"No 'Remove member from group' activities found involving user {user_upn} in the last 30 days.")
    else:
        print(f"Found {len(logs)} log entries:")

        table = []
        csv_logs = []
        group_ids = []
        headers = [Fore.GREEN + "Actor" + Fore.RESET, Fore.YELLOW + "OID" + Fore.RESET, Fore.BLUE + "GDN" + Fore.RESET,
                   Fore.WHITE + "Date" + Fore.RESET]

        for log in logs:
            initiated_by = log.get('initiatedBy', {})
            user_info = initiated_by.get('user', {})
            user_principal_name = user_info.get('userPrincipalName', 'N/A') if user_info else 'N/A'

            group_object_id = 'N/A'
            group_display_name = 'N/A'

            for resource in log.get('targetResources', []):
                if resource.get('type') == 'User' and resource.get('userPrincipalName') == user_upn:
                    for modified_property in resource.get('modifiedProperties', []):
                        if modified_property.get('displayName') == 'Group.ObjectID':
                            group_object_id = modified_property.get('oldValue', 'N/A').strip('"')
                            group_ids.append(group_object_id)
                        if modified_property.get('displayName') == 'Group.DisplayName':
                            group_display_name = modified_property.get('oldValue', 'N/A').strip('"')

            table.append([
                Fore.GREEN + user_principal_name + Fore.RESET + ' ' + Fore.RED + 'removed' + Fore.RESET,
                Fore.YELLOW + group_object_id + Fore.RESET,
                Fore.BLUE + group_display_name + Fore.RESET,
                Fore.WHITE + log['activityDateTime'] + Fore.RESET
            ])
            csv_logs.append(
                [user_principal_name + ' removed', group_object_id, group_display_name, log['activityDateTime']])

        print(tabulate(table, headers=headers, tablefmt="grid"))

        export_choice = input("Do you want to export this data to a CSV file? (yes/no): ").strip().lower()
        if export_choice == 'yes':
            filepath = "C:/audit_logs.csv"
            export_to_csv(csv_logs, filepath)
            print(f"Data exported to {filepath}")

        re_add_choice = input("Do you want to re-add the user to all these groups? (yes/no): ").strip().lower()
        if re_add_choice == 'yes':
            re_add_user_to_groups(access_token, user_upn, group_ids)


if __name__ == "__main__":
    main()