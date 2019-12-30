#! python3
# Gets calendar data from Google Calendar API and writes to a CSV file

from __future__ import print_function
import datetime as dt
import pickle
import os.path
import sys
import csv
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']


def check_creds(credPath):
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
                credPath + 'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    return service


def get_cal_ids(service):
    # Return all calendar names (key) and ids (value) associated with the master calendar email address
    page_token = None
    calids = {}
    while True:
        calendar_list = service.calendarList().list(pageToken=page_token).execute()
        for calendar_list_entry in calendar_list['items']:
            calids[calendar_list_entry['summary']] = calendar_list_entry['id']
        page_token = calendar_list.get('nextPageToken')
        if not page_token:
            return calids

            
def format_isodate(dateToFormat):
    # Removes decimal seconds and converts to date
    dateToFormat = str(dateToFormat).split('.')[0]
    dateToFormat = (dt.datetime.strptime(dateToFormat, '%Y-%m-%d %H:%M:%S')).isoformat() + 'Z'
    return dateToFormat


def calendar_data(calName, service, calids, dateMin = dt.datetime.utcnow().isoformat() + 'Z'):  # default, .now(), 'Z' indicates utc standard time
    # Call the Calendar API
    # Get the id from the calendar name
    for cal in calids:
        if cal == calName:
            calid = calids.get(cal)
            print(f'Loading events from calendar {cal}...')

    events_result = service.events().list(calendarId= calid, timeMin=dateMin,
                                        maxResults=60, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])
    calData = []
    eventData = []

    if not events:
        print('No upcoming events found.')

    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        end = event['end'].get('dateTime', event['end'].get('date'))
        title = event['summary']
         # Split time from date
        try:
            startDate, startTime = start.split('T')
            startTime = startTime.split('-')[0]
        except:
            startDate = start
            startTime = 'All day'
        try: 
            endDate, endTime = end.split('T')
            endTime = endTime.split('-')[0]
        except:
            endDate = end
            endTime = 'No end time'
        try:
            details = event['description']
        except:
            details = ''
   
        eventData = [startDate, startTime, endDate, endTime, title, details, calName]
        calData.append(eventData)
    return calData


def sort_events(dMin, dMax, calList, service, calid, dateThreshold = format_isodate(dt.datetime.now())):
    # Returns a list of event data within a set date range
    calEvents = []
    print(f'Date range : {dMin.split("T")[0]} - {dMax.split("T")[0]}')

    # Convert dates for comparison
    dtMax = dt.datetime.strptime(dMax, '%Y-%m-%dT%H:%M:%SZ')
    dateThreshold = dt.datetime.strptime(dateThreshold, '%Y-%m-%dT%H:%M:%SZ')

    # Get data from the calendars
    for calendar in calList:
        calData = calendar_data(calendar, service, calid, dMin)

        # Create relevant calendar data
        for i in range(0, len(calData)):
            # Convert str date to date for range check
            startDateAsDate = dt.datetime.strptime(calData[i][0], '%Y-%m-%d')
            endDateAsDate = dt.datetime.strptime(calData[i][2], '%Y-%m-%d')
            if endDateAsDate < dateThreshold:
                continue
            if startDateAsDate > dtMax:
                break

            # New list of events that meet the date range
            calEvents.append(calData[i])

    # Sort the calendar data by date
    calEvents.sort()
    return calEvents


def write_csv(fileName, dataList):
    # Write 
    with open(fileName, 'w') as csvfile:
        csvWriter = csv.writer(csvfile, dialect='excel')
        print(f'Writing calendar data to {fileName}...')
        for csvrow in dataList:
            csvWriter.writerow(csvrow)
        csvfile.close()


# ==== MAIN ====

# Ensure proper calendar and credentials
# TODO function that finds the dropbox path
try:
    dropBoxPath = sys.argv[4]
except:
    sys.exit('Please designate a dropbox path in argument 4')

credPath = dropBoxPath + '\\PythonScripts\\'
service = check_creds(credPath)

# Get the names and ids of all calendars associated with master calendar the email address
calid = get_cal_ids(service)

# Pull data from the following calendars
calendars = [
    'village.ballroom@gmail.com',
    'Pub Reservations',
    'COTD Events Calendar',
    'Pub Events'
]

# Date range thresholds from arguments
dateMin = dt.datetime.strptime(sys.argv[1], '%m/%d/%Y').isoformat() + 'Z'
dateMax = dt.datetime.strptime(sys.argv[2], '%m/%d/%Y').isoformat() + 'Z'

# Get event data within range
calEvents = sort_events(dateMin, dateMax, calendars, service, calid)

# Get VACATIONS data and add to final list
dMin = format_isodate(dt.datetime.utcnow() + dt.timedelta(-14)) # Now minus n days
calendars = ['VACATIONS']
calEvents = calEvents + sort_events(dMin, dateMax, calendars, service, calid, dateMin)

# Write calEvents to csv file
try:
    targetFilePath = sys.argv[3]
except:
    targetFilePath = os.getcwd()

write_csv(targetFilePath + '\\oph_sch_cal_data.csv', calEvents)

# Create a txt file to signal VBA that the script is done running.  VBA will delete the file
with open(targetFilePath + '\\done.txt', 'w') as doneStatus:
    doneStatus.write('')

print('Done.')