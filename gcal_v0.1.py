# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 23:16:28 2019

@author: Andrej
"""

from __future__ import print_function
import pickle
import os.path
import sys
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import xls_read

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

json = resource_path('credentials.json')

def main():
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
                json, SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)
    
    
    #Naredi seznam šihtov iz xls_read modula
    seznam = xls_read.seznam
    
    for day,time in enumerate(seznam):
        if (time == (15,00,21,00) or time == (15,00,23,00) or
        time == (15,00,20,00) or time == (8,00,15,00) or
        time == (8,30,15,00) or time == (9, 0, 15, 0) or
        time == (8,00,13,00) or time == (8,00,17,00) or
        time == (13,00,21,00) or time == (13,00,21,00) or
        time == (9,00,20,00)):
            h1,m1,h2,m2 = time
            month = xls_read.month
            event = {
  'summary': 'Atlantis',
  'start': {
    'dateTime': '2019-{}-{}T{}:{}:00'.format(month,day,h1,m1),
    'timeZone': 'Europe/Ljubljana',
  },
  'end': {
    'dateTime': '2019-{}-{}T{}:{}:00'.format(month,day,h2,m2),
    'timeZone': 'Europe/Ljubljana',
  },
  'reminders': {
    'useDefault': False,
    'overrides': [
      {'method': 'popup', 'minutes': 30},
    ],
  },
}

            event = service.events().insert(calendarId='primary', body=event).execute()
            print('Opravljeno: {}.{}: od {}.{} do {}.{}: OK'.format(day,month,h1,m1,h2,m2))

        else:
            pass
    print('© Andrej Sovič\nMaj 2019\nV primeru težav se obrni na asovic@me.com')
    input()
if __name__ == '__main__':
    main()