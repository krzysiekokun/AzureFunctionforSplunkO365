import requests
import json
from datetime import datetime

client_id = 'f3dc1fe5-d774-4ac0-960f-67138f65a42d'
client_secret = '7H_8Q~xfYahdoMGiSdqNGwAKfsMA2J0Q8tgP2bRg'
tenant_id = '9be8208b-1b85-4ac6-ac80-bbe2d6889d37'
users_file = 'users.txt'
output_file = 'all_events.json'

def get_access_token(client_id, client_secret, tenant_id):
    url = f'https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token'
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(url, headers=headers, data=data)
    if response.status_code == 200:
        return response.json()['access_token']
    else:
        print("Error:", response.status_code, response.text)
        return None

def get_users(access_token, users_file):
    users = []
    with open(users_file, 'r', encoding='utf-8-sig') as file:
        for line in file:
            user_identifier = line.strip()
            url = f'https://graph.microsoft.com/v1.0/users/{user_identifier}'
            headers = {
                'Authorization': 'Bearer ' + access_token,
                'Content-Type': 'application/json',
                'Accept-Charset': 'utf-8'
            }
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                users.append(response.json())
            else:
                print("Error:", response.status_code, response.text)
    return users

def get_all_events(access_token, user_id):
    events = []
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/events'
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
        'Accept-Charset': 'utf-8-sig'
    }

    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            json_response = response.json()
            events.extend(json_response['value'])

            # Sprawdzanie czy istnieje kolejna strona wyników
            url = json_response.get('@odata.nextLink', None)
        else:
            print("Error:", response.status_code, response.text)
            return []

    return events

all_users_events = {}
access_token = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token, users_file)

    for user in users:
        user_id = user['id']
        user_email = user['mail']

        user_events = get_all_events(access_token, user_id)
        all_users_events[user_email] = user_events

    with open(output_file, 'w', encoding='utf-8-sig') as file:
        json.dump(all_users_events, file, ensure_ascii=False, indent=2)

    print(f'Zapisano pełne dane wszystkich wydarzeń do pliku {output_file}')
else:
    print("Nie można uzyskać tokena dostępu.")