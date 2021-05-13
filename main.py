from __future__ import print_function
import datetime
import os.path
import argparse
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import requests, json
from google.oauth2.credentials import Credentials
from datetime import datetime, timedelta
from difflib import SequenceMatcher

SCOPES = ['https://www.googleapis.com/auth/calendar']

class CalendarNotFoundException(Exception):
    pass

class CalendarChecker:
    """
    UniBo calendar checking tool
    """

    def __init__(self):
        """
        Inits the script, checking credentials, args, available calendars.
        """
        # If modifying these scopes, delete the file token.pickle.
        self.parse_args()
        self.creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            self.creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.args["credentials"], SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(self.creds.to_json())

        self.service = build('calendar', 'v3', credentials=self.creds)
        self.get_calendars()
        self.choose_calendar()

    def parse_args(self):
        """Parses the given arguments"""
        parser = argparse.ArgumentParser()
        parser.add_argument('-c', '--calendar', dest='calendar_name', type=str, help="Destination calendar name",
                            required=True)
        parser.add_argument('-t', '--credentials', dest='credentials', type=str, help="Google credentials file",
                            required=True)
        parser.add_argument('-s', '--start', dest='start', type=str, help="Starting insert date, %d-%m-%Y",
                            required=False)
        parser.add_argument('-e', '--end', dest='end', type=str, help="Ending insert date, %d-%m-%Y", required=False)
        self.args = parser.parse_args()
        if self.args["start"]:
            self.fromDate = datetime.datetime.strptime(self.args["start"], "%d-%m-%Y")
        else:
            self.fromDate = datetime.today()
        if self.args["end"]:
            self.toDate = datetime.datetime.strptime(self.args["end"], "%d-%m-%Y")
        else:
            self.toDate = self.fromDate + timedelta(days=5)  # By default, we'll use a timedelta of 5 days

    def insert_events(self, url, notExams=[]):
        """
        Inserts the events in the calendar
        :param url: url of the UniBo calendar
        :param notExams: exams to skip in calendar adding
        :return:
        """
        response = json.loads(requests.get(url).text)
        i = 0
        eventDate = datetime.datetime.strptime(response[i]['start'], "%Y-%m-%dT%H:%M:%S")
        while eventDate < self.fromDate:
            i += 1
            eventDate = datetime.datetime.strptime(response[i]['start'], "%Y-%m-%dT%H:%M:%S")

        while eventDate <= self.toDate:
            jEvent = response[i]
            if jEvent['cod_modulo'] not in notExams:
                location = ''
                desc = ''
                if (len(jEvent['aule']) > 0):
                    location = jEvent['aule'][0]['des_indirizzo'].replace(' -', ',')
                    for a in jEvent['aule']:
                        desc += a['des_risorsa'] + ', ' + a['des_piano'] + ' - ' + a['des_ubicazione'] + '\n'
                desc += 'Professor: ' + jEvent['docente']
                if type(jEvent['teams']) is str:
                    desc += '\nTeams: ' + jEvent['teams'] + '\n'
                event = {
                    'summary': jEvent['cod_modulo'] + ' - ' + jEvent['title'],
                    'location': location,
                    'description': desc,
                    'start': {
                        'dateTime': jEvent['start'],
                        'timeZone': 'Europe/Rome',
                    },
                    'end': {
                        'dateTime': jEvent['end'],
                        'timeZone': 'Europe/Rome',
                    },
                    'recurrence': [
                        # 'RRULE:FREQ=DAILY;COUNT=2'
                    ],
                    'attendees': [
                        # {'email': 'lpage@example.com'},
                        # {'email': 'sbrin@example.com'},
                    ],
                    'reminders': {
                        'useDefault': False,
                        'overrides': [
                            {'method': 'popup', 'minutes': 60},
                        ],
                    },
                }
                # if you want to add it to your primary calendar just use calendarId="primary"
                event = self.service.events().insert(calendarId=self.chosen_calendar,
                                                     body=event).execute()
                print('Event created succesfully : %s' % (event.get('htmlLink')))
            i += 1
            eventDate = datetime.datetime.strptime(response[i]['start'], "%Y-%m-%dT%H:%M:%S")

    def choose_calendar(self):
        """
        Checks the found calendars for match of the given calendar name
        :return:
        """
        page_token = None
        self.calendar_list = self.service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in self.calendar_list['items']:
            if similar(calendar_list_entry['summary'], self.args["calendar_name"]) > 0.8:
                self.chosen_calendar = calendar_list_entry['id']
                return
        raise CalendarNotFoundException("No calendar with the provided name was found")


def similar(a, b):
    """
    Checks two strings for a similarity score
    :param a: First string
        :param b: Second string
        :return: Similarity score from 0 to 1
    """
    return SequenceMatcher(None, a, b).ratio()


inserter = CalendarChecker()
inserter.insert_events("https://corsi.unibo.it/2cycle/artificial-intelligence/timetable/@@orario_reale_json?")
