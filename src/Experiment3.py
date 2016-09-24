from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection
# from datetime import datetime
from pytz import timezone

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime
import pickle

print 'cats'
EX_URL = u"https://webmail.olin.edu/EWS/Exchange.asmx"
EX_USERNAME = u'MILKYWAY\\emiller'
EX_PASSWORD = u"[redacted]"
exconnection = None
service = None

# I created this list from
# https://developers.google.com/google-apps/calendar/v3/reference/calendarList/list
# manually.
calendarsToSync = ['primary', 
                   '7mibsonemjp20dsedimtmnrkho@group.calendar.google.com', # OIE
                   '1crj0eos6tonqn7e71mv95m13g@group.calendar.google.com', # Family
                   'a1ugtsnbr7amn62retid68pk0k@group.calendar.google.com' # MedRev
                   ]


def main():
    # readEx()

    gEvents = retrieveAllGoogleEvents()
    print
    return
    for event in gEvents:

        tz = event['start'].get('timeZone')
        start = event['start'].get('dateTime', event['start'].get('date'))
        print tz, start, event['summary']
        copyEventToExchange(event)


def copyEventToExchange(gEvent):
    exRecord = toExchangeRecord(gEvent)
    exRecord.create()


def convertTime(gTime):
    """Converts a RFC3339 or yyyy-mm-dd format
    into a Python datetime that Exchange likes"""

    if 'dateTime' in gTime:
        # Removes redundant Timezone code
        startTime = '-'.join(gTime.get('dateTime').split('-')[:3])
        startTime = datetime.datetime.strptime(
            startTime, '%Y-%m-%dT%H:%M:%S')
        return (timezone(gTime.get('timeZone', 'EST'))
                .localize(startTime))
    else:
        startTime = datetime.datetime.strptime(
            gTime.get('date'), '%Y-%m-%d')
        startTime = startTime.replace(hour=23)
        timezone('UTC').localize(startTime)


def toExchangeRecord(gEvent):

    eEvent = service.calendar().new_event()
    eEvent.subject = gEvent['summary']
    if 'description' in gEvent:
        # Note: it would be cool to encode more information in an HTML body
        eEvent.body = gEvent['description']
    if 'location' in gEvent:
        eEvent.location = gEvent['location']

    # if eEvent.
    if u'dateTime' in gEvent['start']:
        # Event is not all-day
        eEvent.is_all_day = False
        eEvent.start = convertTime(gEvent['start'])
        eEvent.end = convertTime(gEvent['end'])
    else:
        # Event is all-day
        eEvent.is_all_day = True
        eEvent.start = convertTime(gEvent['start'])
        eEvent.end = convertTime(gEvent['end'])

    return eEvent


def connectEx():
    global exconnection, service
    exconnection = ExchangeNTLMAuthConnection(url=EX_URL,
                                              username=EX_USERNAME,
                                              password=EX_PASSWORD)
    service = Exchange2010Service(exconnection)
connectEx()


def getDateRanges():
    return [(timezone("US/Eastern").localize(
                datetime.datetime(year, 1, 1, 1, 0, 0)),
            timezone("US/Eastern").localize(
                datetime.datetime(year + 1, 1, 1, 1, 0, 0)))
            for year in range(2013, 2017)]


def cancelAllEx():
    f = open("backup.pickle", 'w')
    events = readEx()
    pickle.dump(events, f)
    f.close()

    for e in events:
        e.cancel()


def readEx():
    # Set up the connection to Exchange

    global my_calendar
    my_calendar = service.calendar()
    events = []

    for daterange in getDateRanges():
        events.extend(my_calendar.list_events(
            start=daterange[0],
            end=daterange[1],
            details=True
        ).events)

    for event in events:
        print "{start} {stop} - {subject}".format(
            start=event.start,
            stop=event.end,
            subject=event.subject
        )
    return events


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
        print 'Storing credentials to ' + credential_path
    return credentials


def retrieveAllGoogleEvents():
    global allGoogleEvents
    allGoogleEvents = []
    for cal in calendarsToSync:
        allGoogleEvents.extend(retrieveGoogleEvents(calendarId=cal))
    return allGoogleEvents


def retrieveGoogleEvents(calendarId='primary', verbose=False):
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    global gservice
    gservice = discovery.build('calendar', 'v3', http=http)

    # 'Z' indicates UTC time
    # now = datetime.datetime.utcnow().isoformat() + 'Z'
    searchstart = datetime.datetime(2010, 1, 1).isoformat() + 'Z'
    searchend = datetime.datetime(2020, 1, 1).isoformat() + 'Z'

    calName = gservice.calendars().get(
        calendarId=calendarId).execute()['summary']
    print '\n\nRetrieving events from', calName
    global googleEvents
    googleEvents = []

    try:
        page_token = None
        while True:
            eventsResult = gservice.events().list(
                calendarId=calendarId, timeMin=searchstart, singleEvents=True,
                timeMax=searchend, pageToken=page_token,  # maxResults=100,
                orderBy='startTime').execute()
            # calendar_list = service.calendarList().list(
            #     pageToken=page_token).execute()
            googleEvents.extend(eventsResult['items'])
            print len(googleEvents), 'events recieved from', calName
            # if len(googleEvents)>20:
            #     break
            page_token = eventsResult.get('nextPageToken')
            if not page_token:
                break

    except client.AccessTokenRefreshError:
        print ('The credentials have been revoked or expired, please re-run'
               'the application to re-authorize.')

    if verbose:
        if not googleEvents:
            print 'No upcoming events found.'
        for event in googleEvents:
            tz = event['start'].get('timeZone')
            start = event['start'].get('dateTime', event['start'].get('date'))
            print tz, start, event['summary']

    return googleEvents


if __name__ == '__main__':
    main()
