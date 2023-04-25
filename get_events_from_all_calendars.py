import requests
import csv
from datetime import datetime

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

def get_all_events(access_token, user_id, start_time, end_time):
    start_time = start_time.isoformat()
    end_time = end_time.isoformat()

    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/calendarview?startDateTime={start_time}&endDateTime={end_time}'
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
        
def filter_recurring_events(events, user_email, start_time, end_time):
    filtered_events = []
    for event in events:
        event_start_time = datetime.fromisoformat(event['start']['dateTime'])
        event_end_time = datetime.fromisoformat(event['end']['dateTime'])

        if ('recurrence' in event and
            event['organizer']['emailAddress']['address'] == user_email and
            event_start_time >= start_time and event_end_time <= end_time):
            filtered_events.append(event)
    return filtered_events





def save_to_csv(recurring_events, output_file):
    with open(output_file, 'a', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file, delimiter=';')
        for event in recurring_events:
            organizer = event['organizer']['emailAddress']['name']
            subject = event['subject']
            attendees_count = len(event['attendees'])
            recurrence_pattern = event['recurrence']['pattern']['type'] if event['recurrence'] and 'pattern' in event['recurrence'] else 'N/A'
            is_cancelled = event['isCancelled']
            event_create_date = event['createdDateTime']
            recurrence_range_start_date = event['recurrence']['range']['startDate'] if event['recurrence'] and 'range' in event['recurrence'] else 'N/A'
            recurrence_range_end_date = event['recurrence']['range']['endDate'] if event['recurrence'] and 'range' in event['recurrence'] else 'N/A'
            start_time = event['start']['dateTime']
            end_time = event['end']['dateTime']

            csv_writer.writerow([
                organizer,
                subject,
                attendees_count,
                recurrence_pattern,
                start_time,
                end_time,
                is_cancelled,
                event_create_date,
                recurrence_range_start_date,
                recurrence_range_end_date
            ])

# Dodaj nagłówek do pliku CSV przed rozpoczęciem zapisywania danych
with open(output_file, 'w', newline='', encoding='utf-8') as file:
    csv_writer = csv.writer(file, delimiter=';')
    csv_writer.writerow(['Organizer', 'Event Subject', 'Attendees Count', 'Recurrence Pattern', 'Start Time', 'End Time', 'Is Cancelled', 'Event Create Date', 'Recurrence Range Start Date', 'Recurrence Range End Date'])

access_token = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token, users_file)

    start_time = datetime(2023, 1, 1)
    end_time = datetime(2023, 6, 30)

    for user in users:
        user_id = user['id']
        user_email = user['mail']
        
        all_events = get_all_events(access_token, user_id, start_time, end_time)
        recurring_events = filter_recurring_events(all_events, user_email, start_time, end_time)

        save_to_csv(recurring_events, output_file)

    print(f'Zapisano cykliczne spotkania do pliku {output_file}')
else:
    print("Nie można uzyskać tokena dostępu.")
