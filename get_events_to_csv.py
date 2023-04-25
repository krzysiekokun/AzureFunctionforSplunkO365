import requests
import csv
from datetime import datetime

client_id = 'f3dc1fe5-d774-4ac0-960f-67138f65a42d'
client_secret = '7H_8Q~xfYahdoMGiSdqNGwAKfsMA2J0Q8tgP2bRg'
tenant_id = '9be8208b-1b85-4ac6-ac80-bbe2d6889d37'
users_file = 'users.txt'
output_file = 'filtered_events.csv'

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

def save_to_csv(user_email, events, output_file):
    with open(output_file, 'a', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file, delimiter=';')
        for event in events:
            organizer_name = event['organizer']['emailAddress']['name']
            subject = event['subject']
            attendees_count = len(event['attendees'])
            is_organizer = user_email == event['organizer']['emailAddress']['address']
            recurrence = event.get('recurrence', None) is not None

            csv_writer.writerow([
                organizer_name,
                subject,
                attendees_count,
                is_organizer,
                recurrence
            ])

# Dodaj nagłówek do pliku CSV przed rozpoczęciem zapisywania danych
with open(output_file, 'w', newline='', encoding='utf-8') as file:
    csv_writer = csv.writer(file, delimiter=';')
    csv_writer.writerow(['Organizer Name', 'Subject', 'Attendees Count', 'Is Organizer', 'Has Recurrence'])

access_token = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token, users_file)

    for user in users:
        user_id = user['id']
        user_email = user['mail']

        user_events = get_all_events(access_token, user_id)

        save_to_csv(user_email, user_events, output_file)

    print(f'Zapisano wybrane dane wydarzeń do pliku {output_file}')
else:
    print("Nie można uzyskać tokena dostępu.")