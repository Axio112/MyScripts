import msal
import requests

# Azure AD app details
CLIENT_ID = '----------------------------'
CLIENT_SECRET = '----------------------------'
TENANT_ID = '----------------------------'

# Authentication endpoint
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Microsoft Graph API endpoint
GRAPH_API_ENDPOINT = 'https://graph.microsoft.com/v1.0/'

# Scope
SCOPE = ["https://graph.microsoft.com/.default"]

# Relevant SKU IDs
OFFICE365_E3_SKU_ID = '05e9a617-0261-4cee-bb44-138d3ef5d965'
DEFENDER_SKU_ID = '4ef96642-f096-40de-a3e9-d83fb2f90211'
BUSINESS_PREMIUM_SKU_ID = '094e7854-93fc-4d55-b2c0-3ab5369ebdc1'


def get_access_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    token_response = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" in token_response:
        print("Access token obtained successfully")
        return token_response['access_token']
    else:
        print(f"Could not obtain access token: {token_response}")
        raise Exception("Could not obtain access token")


def get_all_users(access_token):
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }
    users = []
    url = f"{GRAPH_API_ENDPOINT}users?$select=id,displayName,userPrincipalName,assignedLicenses"
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            result = response.json()
            users.extend(result.get('value', []))
            url = result.get('@odata.nextLink', None)
        else:
            print(f"API call failed with status code {response.status_code} and response: {response.text}")
            raise Exception(f"API call failed with status code {response.status_code}")
    return users


def filter_users(users):
    e3_not_defender_users = []
    defender_and_business_premium_users = []

    for user in users:
        has_e3 = False
        has_defender = False
        has_business_premium = False
        if 'assignedLicenses' in user and user['assignedLicenses']:
            for license in user['assignedLicenses']:
                if license['skuId'] == OFFICE365_E3_SKU_ID:
                    has_e3 = True
                if license['skuId'] == DEFENDER_SKU_ID:
                    has_defender = True
                if license['skuId'] == BUSINESS_PREMIUM_SKU_ID:
                    has_business_premium = True
            if has_e3 and not has_defender:
                e3_not_defender_users.append(user)
            if has_defender and has_business_premium:
                defender_and_business_premium_users.append(user)

    return e3_not_defender_users, defender_and_business_premium_users


def main():
    try:
        access_token = get_access_token()

        # Get all users
        users = get_all_users(access_token)

        # Filter users
        e3_not_defender_users, defender_and_business_premium_users = filter_users(users)

        print("Users with Office 365 E3 but not Microsoft Defender for Office 365 (Plan 1):")
        for user in e3_not_defender_users:
            print(f"  - {user['displayName']} ({user['userPrincipalName']})")

        print("\nUsers with Microsoft Defender for Office 365 (Plan 1) and Business Premium:")
        for user in defender_and_business_premium_users:
            print(f"  - {user['displayName']} ({user['userPrincipalName']})")
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()