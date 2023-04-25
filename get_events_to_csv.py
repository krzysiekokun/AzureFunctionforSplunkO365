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
        'Accept-Charset': 'utf-8-sig'
    }
    
    events = []
    while url:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            result = response.json()
            events.extend(result['value'])
            url = result.get('@odata.nextLink', None)
        else:
            print("Error:", response.status_code, response.text)
            return []
    
    return events

def recurrence_pattern_to_string(pattern_type, pattern_interval, pattern_days_of_week, range_type):
    return f"{pattern_type} (Interval: {pattern_interval}, Days of Week: {pattern_days_of_week}, Range Type: {range_type})"

def save_to_csv(user_email, events, output_file):
    with open(output_file, 'a', newline='', encoding='utf-8') as file:
        csv_writer = csv.writer(file, delimiter=';')
        for event in events:
            recurrence = event.get('recurrence', None)
            if recurrence is not None:
                pattern_type = recurrence['pattern']['type']
                pattern_interval = recurrence['pattern']['interval']
                pattern_days_of_week = recurrence['pattern'].get('daysOfWeek', '')
                range_type = recurrence['range']['type']

                recurrence_pattern_str = recurrence_pattern_to_string(pattern_type, pattern_interval, pattern_days_of_week, range_type)

                organizer_name = event['organizer']['emailAddress']['name']
                subject = event['subject']
                attendees_count = len(event['attendees'])
                is_organizer = user_email == event['organizer']['emailAddress']['address']
                created_date_time = event['createdDateTime']
                last_modified_date_time = event['lastModifiedDateTime']
                start_date_time = event['start']['dateTime']
                end_date_time = event['end']['dateTime']
                is_cancelled = event['isCancelled']
                is_all_day = event['isAllDay']
                importance = event['importance']

                csv_writer.writerow([
                    organizer_name,
                    subject,
                    attendees_count,
                    is_organizer,
                    recurrence_pattern_str,
                    created_date_time,
                    last_modified_date_time,
                    start_date_time,
                    end_date_time,
                    is_cancelled,
                    is_all_day,
                    importance
                ])

with open(output_file, 'w', newline='', encoding='utf-8-sig') as file:
    csv_writer = csv.writer(file, delimiter=';')
    csv_writer.writerow(['Organizer Name', 'Subject', 'Attendees Count', 'Is Organizer', 'Recurrence Pattern', 'Created Date Time', 'Last Modified Date Time', 'Start Date Time', 'End Date Time', 'Is Cancelled', 'Is All Day', 'Importance'])

access_token = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token, users_file)

    for user in users:
        user_id = user['id']
        user_email = user['mail']

        print(f'Processing events for user: {user_email}')
        user_events = get_all_events(access_token, user_id)

        save_to_csv(user_email, user_events, output_file)

        recurring_events_count = sum(1 for event in user_events if event.get('recurrence', None) is not None)
        print(f'Total events for user {user_email}: {len(user_events)} (recurring events: {recurring_events_count})')

    print(f'Saved selected event data to file {output_file}')
else:
    print("Cannot obtain access token.")