import requests

# Uzupełnij te wartości swoimi danymi
client_id = 'f3dc1fe5-d774-4ac0-960f-67138f65a42d'
client_secret = '7H_8Q~xfYahdoMGiSdqNGwAKfsMA2J0Q8tgP2bRg'
tenant_id = '9be8208b-1b85-4ac6-ac80-bbe2d6889d37'

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

def get_users(access_token):
    url = 'https://graph.microsoft.com/v1.0/users'
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['value']
    else:
        print("Error:", response.status_code, response.text)
        return []

def get_all_events(access_token, user_id):
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/events'
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['value']
    else:
        print("Error:", response.status_code, response.text)
        return []

def filter_recurring_events(events):
    return [event for event in events if 'recurrence' in event]

access_token = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token)
    for user in users:
        user_id = user['id']
        user_display_name = user['displayName']
        print(f'Cykliczne spotkania użytkownika {user_display_name}:')
        
        all_events = get_all_events(access_token, user_id)
        recurring_events = filter_recurring_events(all_events)
        for event in recurring_events:
            print(event)
        print('\n')
else:
    print("Nie można uzyskać tokena dostępu.")
