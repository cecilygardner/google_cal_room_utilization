from __future__ import print_function, division

from datetime import datetime, timedelta
import argparse
import json
import os
import os.path
import urllib

from requests_oauthlib import OAuth2Session
import asana
import dateutil.parser as date_parser
import pytz


os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '1'

CLIENT_SECRET_FILE = 'client_secret.json'


def gen_parser(utc_now, utc_thirty_days_ago):
    def parse_tz_aware_dt(dt_str):
        return pytz.utc.localize(date_parser.parse(dt_str))

    parser = argparse.ArgumentParser(
        description='Generate statistics for Google Cal room utilization.')
    parser.add_argument(
        '--start-date', type=parse_tz_aware_dt, default=utc_thirty_days_ago,
        help='Start date for the date range to run the report on.')
    parser.add_argument(
        '--end-date', type=parse_tz_aware_dt, default=utc_now,
        help='End date for the date range to run the report on.')
    return parser


def create_google_oauth_session(scopes):
    with open(CLIENT_SECRET_FILE) as client_secret_fobj:
        client_secret_dict = json.load(client_secret_fobj)['installed']

    client_id = client_secret_dict['client_id']
    client_secret = client_secret_dict['client_secret']
    auto_refresh_kwargs = {
        'client_id': client_id,
        'client_secret': client_secret
    }
    auth_uri = client_secret_dict['auth_uri']
    redirect_uri = 'urn:ietf:wg:oauth:2.0:oob'
    token_uri = client_secret_dict['token_uri']

    if os.path.exists('google_oauth_token.json'):
        with open('google_oauth_token.json') as oauth_token_fobj:
            token = json.load(oauth_token_fobj)
        google_oauth_session = OAuth2Session(
            client_id, token=token, auto_refresh_url=token_uri,
            auto_refresh_kwargs=auto_refresh_kwargs,
            token_updater=save_token_as_json)
    else:
        google_oauth_session = OAuth2Session(
            client_id, redirect_uri=redirect_uri, scope=scopes,
            auto_refresh_url=token_uri,
            auto_refresh_kwargs=auto_refresh_kwargs,
            token_updater=save_token_as_json)
        authorization_url, state = google_oauth_session.authorization_url(
            auth_uri)

        print(
            'Please go to {} and authorize access.'.format(authorization_url))
        code = raw_input('Enter the code you recieve: ')

        token = google_oauth_session.fetch_token(
            token_uri, code=code, client_secret=client_secret)

        save_token_as_json(token)

    return google_oauth_session


def save_token_as_json(token):
    with open('google_oauth_token.json', 'w') as oauth_token_fobj:
        json.dump(token, oauth_token_fobj)


def generate_room_util_report(room_list, start_date, end_date):
    room_utilization_strs = []
    google_oauth_session = create_google_oauth_session(
        ['https://www.googleapis.com/auth/calendar.readonly'])

    print('Generating Room Utilization Report...')
    for room in room_list:
        room_name = room['name']
        room_url = room['url']
        room_seats = room['room_seats']
        print('Grabbing data for room {} (URL: {}'.format(room_name, room_url))

        number_of_meetings = 0
        number_of_attendees = 0
        max_capacity = 0

        url = (
            'https://www.googleapis.com/calendar/v3/calendars/{calendar_id}/'
            'events'.format(calendar_id=urllib.quote(room_url)))
        params = {
            'timeMin': start_date.isoformat(),
            'timeMax': end_date.isoformat(),
            'singleEvents': True,
            'orderBy': 'startTime'
        }
        events_response = google_oauth_session.get(
            url, params=params)
        events_response.raise_for_status()

        events = events_response.json().get('items', [])

        if not events:
            print('No events found for room: {}'.format(room_name))
            continue
        for event in events:
            if 'attendees' not in event:
                number_of_attendees += 1
            else:
                number_of_attendees += len(
                    [attendee for attendee in event['attendees'] if
                     attendee['responseStatus'] != 'declined'])

        number_of_meetings = len(events)
        max_capacity = float(number_of_meetings * room_seats)
        util = (number_of_attendees / max_capacity) * 100

        room_utilization_strs.append('{} Utilization %: {:0.2f}'.format(
            room_name, util))

    report_string = (
        'The room data below is pulled from calendar events between {} '
        'and {}. \n\nConference room utilization = (number of attendees for '
        'meetings in last 30 days) / (number of meetings * max occupancy of '
        'room). The count of seats includes non-Asana folks '
        "and office hours. If a room has > 100% that means generally it's "
        'over occupied. This seems to be the case for most of the phone rooms '
        'likely due to customers being invited on the calendar invite. The '
        'seat count includes accepted and non-responded calendar '
        'invites.\n\n{}').format(
            start_date.strftime("%m/%d/%Y"), end_date.strftime("%m/%d/%Y"),
            '\n'.join(room_utilization_strs))

    return report_string


def post_all_in_asana_task(asana_config, room_util_report_str, utc_now):
    asana_client = asana.Client.access_token(
        asana_config['personal_access_token'])
    todays_date_str = utc_now.strftime('%m/%d/%Y')
    name = 'Google Calendar Room Utilization Results {}'.format(
        todays_date_str)

    params = {
        'workspace': asana_config['workspace_id'],
        'projects': [asana_config['project_id']],
        'name': name,
        'notes': room_util_report_str
    }

    asana_client.tasks.create(params)
    print("Utilization task created in Asana workspace")


def read_room_list_from_json():
    with open('rooms.json') as rooms_fobj:
        return json.load(rooms_fobj)['rooms']


def read_config():
    with open('asana_config.json') as config_fobj:
        return json.load(config_fobj)


def main():
    utc_now = pytz.utc.localize(datetime.utcnow())
    utc_thirty_days_ago = utc_now - timedelta(days=30)
    parser = gen_parser(utc_now, utc_thirty_days_ago)
    args = parser.parse_args()
    room_list = read_room_list_from_json()
    room_util_report_str = generate_room_util_report(
        room_list, args.start_date, args.end_date)
    asana_config = read_config()
    post_all_in_asana_task(asana_config, room_util_report_str, utc_now)
    print("YOU DID IT! :)")

if __name__ == '__main__':
    main()
