import requests
import csv

# Uzupełnij te wartości swoimi danymi
client_id = 'f3dc1fe5-d774-4ac0-960f-67138f65a42d'
client_secret = '7H_8Q~xfYahdoMGiSdqNGwAKfsMA2J0Q8tgP2bRg'
tenant_id = '9be8208b-1b85-4ac6-ac80-bbe2d6889d37'
users_file = 'users.txt'
output_file = 'recurring_events.csv'

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
    with open(users_file, 'r', encoding='utf-8') as file:
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
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/events'
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
        'Accept-Charset': 'utf-8'
    }
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        return response.json()['value']
    else:
        print("Error:", response.status_code, response.text)
        return []

def filter_recurring_events(events):
    return [event for event in events if 'recurrence' in event]

def save_to_csv(recurring_events, output_file):
    with open(output_file, 'w', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file)
        csv_writer.writerow(['User', 'Event Subject', 'Start Time', 'End Time', 'Recurrence Pattern'])
        for user_events in recurring_events:
            user_name = user_events['user_name']
            for event in user_events['events']:
                recurrence_pattern = event['recurrence']['pattern'] if event['recurrence'] and 'pattern' in event['recurrence'] else 'N/A'
                csv_writer.writerow([
                    user_name,
                    event['subject'],
                    event['start']['dateTime'],
                    event['end']['dateTime'],
                    recurrence_pattern
                ])

access_token = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token, users_file)
    all_recurring_events = []
    for user in users:
        user_id = user['id']
        user_display_name = user['displayName']
        
        all_events = get_all_events(access_token, user_id)
        recurring_events = filter_recurring_events(all_events)
        all_recurring_events.append({'user_name': user_display_name, 'events': recurring_events})

    save_to_csv(all_recurring_events, output_file)
    print(f'Zapisano cykliczne spotkania do pliku {output_file}')
else:
    print("Nie można uzyskać tokena dostępu.")

