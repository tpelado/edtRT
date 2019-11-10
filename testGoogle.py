from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import json


def pprint(jsn):
    print(json.dumps(jsn, indent=4, sort_keys=True))


def createEventObject(summary, location, description, start, end):
    event = {
    'summary': summary,
    'location': location,
    'description': description,
    'start': {
        'dateTime': start,
        'timeZone': 'Europe/Paris',
    },
    'end': {
        'dateTime': end,
        'timeZone': 'Europe/Paris',
    },
    'reminders': {
        'useDefault': False,
    },
    }

    return event



#id du BON CALENDAR
idC = "7auhv9oniguke3igbpmuuqlv9c@group.calendar.google.com"

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def addToGcal():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    calendriers = service.calendarList().list().execute()




    eve = createEventObject("Test Object", "Test","This is a test",
    (datetime.datetime.now()).replace(microsecond=0).isoformat(),(datetime.datetime.now()+datetime.timedelta(hours=2)).replace(microsecond=0).isoformat())
    pprint(eve)


    test = service.events().insert(calendarId=idC,body=eve).execute()


    # events_result = service.events().list(calendarId='primary', timeMin=now,
    #                                     maxResults=10, singleEvents=True,
    #                                     orderBy='startTime').execute()
    # events = events_result.get('items', [])
    #
    # if not events:
    #     print('No upcoming events found.')
    # for event in events:
    #     start = event['start'].get('dateTime', event['start'].get('date'))
    #     print(start, event['summary'])

if __name__ == '__main__':
    main()