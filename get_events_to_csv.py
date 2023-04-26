import requests
import csv
from datetime import datetime
from dateutil.relativedelta import relativedelta
import dateutil.parser
from colorama import Fore, Style, init

# Initialize colorama
init(autoreset=True)

client_id = 'f3dc1fe5-d774-4ac0-960f-67138f65a42d'
client_secret = '7H_8Q~xfYahdoMGiSdqNGwAKfsMA2J0Q8tgP2bRg'
tenant_id = '9be8208b-1b85-4ac6-ac80-bbe2d6889d37'
users_file = 'users.txt'
output_file = 'filtered_events.csv'

start_year = 2023
start_month = 1
end_year = 2023
end_month = 6

start_date = dateutil.parser.parse(f"{start_year}-{start_month:02d}-01T00:00:00Z")
end_date = dateutil.parser.parse(f"{end_year}-{end_month:02d}-01T00:00:00Z")

date_ranges = []
while start_date < end_date:
    range_start = start_date
    range_end = start_date + relativedelta(months=1) - relativedelta(seconds=1)
    date_ranges.append((range_start.isoformat(), range_end.isoformat()))
    start_date = range_end + relativedelta(seconds=1)

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
        result = response.json()
        return result['access_token'], datetime.now(), result['expires_in']
    else:
        print("Error:", response.status_code, response.text)
        return None, None, None

def is_token_expired(token_received_time, token_expires_in):
    elapsed_time = (datetime.now() - token_received_time).seconds
    return elapsed_time >= token_expires_in - 60

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

def get_all_events(access_token, user_id, start_date, end_date):
    url = f'https://graph.microsoft.com/v1.0/users/{user_id}/events'
    headers = {
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
        'Accept-Charset': 'utf-8-sig'
    }
    
    params = {
        '$filter': f"start/dateTime ge '{start_date}' and end/dateTime le '{end_date}'",
        '$orderby': 'start/dateTime'
    }
    
    events = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        if response.status_code == 200:
            result = response.json()
            events.extend(result['value'])
            url = result.get('@odata.nextLink', None)
            if url:
                params = None  # Set params to None to avoid adding the filter again in the next request
        else:
            print("Error:", response.status_code, response.text)
            return []
    
    return events

def save_to_csv(user_email, events, output_file):
    with open(output_file, 'a', newline='', encoding='utf-8-sig') as file:
        csv_writer = csv.writer(file, delimiter=';')
        for event in events:
            organizer_name = event['organizer']['emailAddress']['name']
            subject = event['subject']
            attendees_count = len(event['attendees'])
            is_organizer = event['isOrganizer']
            created_date_time = event['createdDateTime']
            last_modified_date_time = event['lastModifiedDateTime']
            start_date_time = event['start']['dateTime']
            end_date_time = event['end']['dateTime']
            is_cancelled = event['isCancelled']
            is_all_day = event['isAllDay']
            importance = event['importance']

            recurrence = event.get('recurrence', None)
            if recurrence:
                pattern_type = recurrence['pattern']['type']
                pattern_interval = recurrence['pattern']['interval']
                pattern_days_of_week = recurrence['pattern'].get('daysOfWeek', '')
                range_type = recurrence['range']['type']
            else:
                pattern_type = ''
                pattern_interval = ''
                pattern_days_of_week = ''
                range_type = ''

            row = [user_email, organizer_name, subject, attendees_count, is_organizer, pattern_type, pattern_interval, pattern_days_of_week, range_type, created_date_time, last_modified_date_time, start_date_time, end_date_time, is_cancelled, is_all_day, importance]
            csv_writer.writerow(row)

with open(output_file, 'w', newline='', encoding='utf-8-sig') as file:
    csv_writer = csv.writer(file, delimiter=';')
    csv_writer.writerow(['User Email', 'Organizer Name', 'Subject', 'Attendees Count', 'Is Organizer', 'Pattern Type', 'Pattern Interval', 'Pattern Days of Week', 'Range Type', 'Created Date Time', 'Last Modified Date Time', 'Start Date Time', 'End Date Time', 'Is Cancelled', 'Is All Day', 'Importance'])

access_token, token_received_time, token_expires_in = get_access_token(client_id, client_secret, tenant_id)
if access_token:
    users = get_users(access_token, users_file)
    failed_users = []

    total_users = len(users)
    for index, user in enumerate(users, start=1):
        user_id = user['id']
        user_email = user['mail']

        print(Fore.GREEN + f'Processing events for user {index}/{total_users}: {user_email}')
        total_events = 0

        for start_date, end_date in date_ranges:
            if is_token_expired(token_received_time, token_expires_in):
                access_token, token_received_time, token_expires_in = get_access_token(client_id, client_secret, tenant_id)
                if not access_token:
                    print("Failed to refresh access token. Exiting.")
                    break

            try:
                events = get_all_events(access_token, user_id, start_date, end_date)
                save_to_csv(user_email, events, output_file)
                total_events += len(events)
                print(f'Fetched {len(events)} events for user {user_email} between {start_date} and {end_date}')

            except Exception as e:
                print(f"Error processing events for user {user_email}: {e}")
                failed_users.append(user_email)
                break

        print(Fore.RED + f'Finished processing events for user: {user_email}. Total events: {total_events}')

    print(f'Saved selected event data to file {output_file}')
    if failed_users:
        with open('failed_users.txt', 'w', encoding='utf-8') as file:
            for failed_user in failed_users:
                file.write(f"{failed_user}\n")
        print(f"Failed users have been saved to 'failed_users.txt'")
else:
    print("Cannot obtain access token.")



