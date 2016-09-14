#! /usr/bin/python2
"""
This program downloads calendar data from Google Calendar,
then performs a one-way synchronization
to merge that data into a Microsoft Exchange acocunt.

Status: complete prototype, not remotely production-ready.
"""

from pyexchange import Exchange2010Service, ExchangeNTLMAuthConnection
# from datetime import datetime
from pytz import timezone

import httplib2
import os
import time
from unidecode import unidecode

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime
import pickle

__author__ = "Eric Miller"
__email__ = "eric@legoaces.org"

#EX_URL = u"https://webmail.olin.edu/EWS/Exchange.asmx"
EX_URL = u"https://webmail.olin.edu/EWS/Exchange.asmx"
EX_USERNAME = ''
EX_PASSWORD = ''

exservice = None
gservice = None

dateRange = tuple()

def main():
    """This function will implement a complete non-incremental
    one-way synchronization from the Google Calendars to Exchange.
    """

    establishgoogconn()
    loadexpwd()
    global dateRange
    dateRange = getDateRange()
    establishexconn()
    calendarsToSync = getgooglecals()

    global allGoogleEvents  # Testing purposes only. TODO: remove
    allGoogleEvents = []
    for cal in calendarsToSync:
        allGoogleEvents.extend(
            retrieveGoogleEvents(calendarId=cal, verbose=True))

    print '\n{0} events retrieved from {1} calendars'.format(
        len(allGoogleEvents), len(calendarsToSync))

    # Begin to perform synchronization
    print 'Cancelling all Exchange events'
    cancelAllEx()
    print 'Re-adding all Google events to Exchange'
    for event in allGoogleEvents:
        print (unidecode(event.get('summary','No Description Found')))
        copyEventToExchange(event)


def loadexpwd():
    credential_dir = '/.credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'exchange-credentials')
    f = open(credential_path)
    global EX_USERNAME
    global EX_PASSWORD
    EX_USERNAME = f.readline().strip()
    EX_PASSWORD = f.readline().strip()
    f.close()
    

def getgooglecals(verbose=True):
    # Calendar IDs can go here to force them to synchronize,
    # otherwise it is automatically populated from all calendars with "exsync"
    # somewhere in their descriptions
    manualOverrides = []

    allcals = gservice.calendarList().list().execute().get('items')
    calendarsToSync = []

    for cal in allcals:
        if (
                cal.get('description', '').lower().find('exsync') >= 0
                or cal.get('id', '') in manualOverrides):

            calendarsToSync.append(cal)

    if verbose:
        print 'Calendars to synchronize:'
        for c in calendarsToSync:
            print '\t', c.get('summary', 'ERR'), c.get('id', 'ERR')
    return [c.get('id') for c in calendarsToSync]


def copyEventToExchange(gEvent):
    exRecord = toExchangeRecord(gEvent)
    if (exRecord.start.replace(tzinfo=None) > dateRange[0] and
            exRecord.start.replace(tzinfo=None) < dateRange[1]):
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
        return timezone('UTC').localize(startTime)


def toExchangeRecord(gEvent):
    if not exservice:
        print 'Exchange service not initialized'
        establishexconn()

    eEvent = exservice.calendar().new_event()
    eEvent.subject = gEvent['summary']

    description = [
        gEvent.get('description',''),
        datetime.datetime.now().strftime('Last updated: %x - %X')
        ]

    # Note: it would be cool to encode more information in an HTML body
    eEvent.html_body = '\n\n'.join(description)

    
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


def establishexconn():
    global exservice
    exconnection = ExchangeNTLMAuthConnection(url=EX_URL,
                                              username=EX_USERNAME,
                                              password=EX_PASSWORD)
    exservice = Exchange2010Service(exconnection)


def getDateRange():
    return (datetime.datetime.today()+datetime.timedelta(days=-180),
             datetime.datetime.today()+datetime.timedelta(days=180))


def cancelAllEx(verbose = True):
    f = open("backup.pickle", 'w')
    events = readEx()
    pickle.dump(events, f)
    f.close()

    for e in events:
        e.cancel()
        if verbose:
            print "Cancelled:", e.subject


def readEx(verbose=False):
    # Set up the connection to Exchange

    global exchange_calendar
    exchange_calendar = exservice.calendar()
    events = exchange_calendar.list_events(
        start=dateRange[0],
        end=dateRange[1],
        details=True
    ).events
    if verbose:
        print "Got some events!"
    time.sleep(1)

    if verbose:
        for event in events:
            try:
                print "{start} {stop} - {subject}".format(
                    start=event.start,
                    stop=event.end,
                    subject=event.subject
                )
            except:
                print "Error printing event"
    return events


try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None


def goog_get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """

    SCOPES = 'https://www.googleapis.com/auth/calendar.readonly'
    CLIENT_SECRET_FILE = '/.credentials/client_secret.json'
    APPLICATION_NAME = 'Olin calendar synchronization project (beta)'

    credential_dir = '/.credentials'
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'google-calendar-credentials.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME

        credentials = tools.run_flow(flow, store, flags)
        print 'Storing credentials to ' + credential_path
    return credentials


def establishgoogconn():
    """Creates a Google Calendar service object and
    puts in in the global variable gservice

    Returns:
        gservice
    """
    credentials = goog_get_credentials()
    http = credentials.authorize(httplib2.Http())
    global gservice
    gservice = discovery.build('calendar', 'v3', http=http)


def retrieveGoogleEvents(calendarId='primary', verbose=True):
    """Shows basic usage of the Google Calendar API.

    retrieves all of the events from this decade from the calendar
    identified by calendarId, using multiple requests and page tokens
    if needed.
    """

    global gservice
    if not gservice:
        establishgoogconn()

    # 'Z' indicates UTC time
    # now = datetime.datetime.utcnow().isoformat() + 'Z'
    searchstart = datetime.datetime(2010, 1, 1).isoformat() + 'Z'
    searchend = datetime.datetime(2020, 1, 1).isoformat() + 'Z'

    if verbose:
        calName = gservice.calendars().get(
            calendarId=calendarId).execute()['summary']
        print '\n\nRetrieving events from', calName

    global googleEvents  # TODO: remove this. For tests only
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
            if verbose:
                print len(googleEvents), 'events recieved from', calName
            # if len(googleEvents)>20:
            #     break
            page_token = eventsResult.get('nextPageToken')
            if not page_token:
                break

    except client.AccessTokenRefreshError:
        print ('The credentials have been revoked or expired, please re-run'
               'the application to re-authorize.')

    # if verbose:
    #     if not googleEvents:
    #         print 'No upcoming events found.'
    #     for event in googleEvents:
    #         tz = event['start'].get('timeZone')
    #         start = event['start'].get('dateTime',
    #             event['start'].get('date'))
    #         print tz, start, event['summary']

    return googleEvents


if __name__ == '__main__':
    main()
