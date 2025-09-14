from __future__ import print_function
import datetime
import pickle
import os.path
import tkinter as tk
from tkinter import messagebox
from google.oauth2.credentials import Credentials
from google.auth.exceptions import RefreshError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build


SCOPES = ['https://www.googleapis.com/auth/calendar']

def get_calendar_service():
    creds = None

    if os.path.exists('token.json'):
        try:
            with open('token.json', 'rb') as token:
                creds = pickle.load(token)
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
        except (RefreshError, Exception) as e:
            print("🔴 Error refreshing token:", e)
            os.remove('token.json')
            show_auth_error_popup("Your Google Calendar access has expired or been revoked. Please sign in again.")

    if not creds or not creds.valid:
        try:
            if not os.path.exists("credentials.json"):
                print("🚨 credentials.json missing in current directory!")
                show_auth_error_popup("Missing 'credentials.json'. Please add your Google API credentials.")
                return None

            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

            with open('token.json', 'wb') as token:
                pickle.dump(creds, token)
        except Exception as e:
            print("🚨 Failed during OAuth flow:", e)
            show_auth_error_popup(f"Google Calendar authorization failed.\n\n{str(e)}")
            return None

    try:
        return build('calendar', 'v3', credentials=creds)
    except Exception as e:
        print("🚨 Failed to build calendar service:", e)
        show_auth_error_popup(f"Failed to connect to Google Calendar service.\n\n{str(e)}")
        return None

def show_auth_error_popup(message):
    root = tk.Tk()
    root.withdraw()  # Hide root window
    messagebox.showwarning("Google Calendar Error", message)
    root.destroy()

def list_upcoming_events(n=5):
    service = get_calendar_service()
    now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          maxResults=n, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    event_list = []
    if not events:
        event_list.append('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        event_list.append(f"{start}: {event['summary']}")
    return event_list

def create_event(summary, start_time, end_time):
    service = get_calendar_service()
    event = {
        'summary': summary,
        'start': {
            'dateTime': start_time,
            'timeZone': 'Asia/Kolkata',
        },
        'end': {
            'dateTime': end_time,
            'timeZone': 'Asia/Kolkata',
        },
    }
    event = service.events().insert(calendarId='primary', body=event).execute()
    return event.get('htmlLink')
