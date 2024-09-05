import datetime as dt
import dateutil
import os
import json
import pickle
import pytz
import time

from bs4 import BeautifulSoup
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

import outlook
import config

SCOPES = ["https://www.googleapis.com/auth/calendar.events"]


def authenticate_google():
    creds = None
    # authenticate google api credentials
    if os.path.exists(config.google_token_path):
        creds = Credentials.from_authorized_user_file(config.google_token_path, SCOPES)
    # if not valid creds available, let user log in
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                    config.google_creds_path, SCOPES)#, access_type='offline')
            creds = flow.run_local_server(port=0)
        with open(config.google_token_path, 'w') as token:
            token.write(creds.to_json())

    service = build("calendar", "v3", credentials=creds)
    se = service.events()

    print("Authenticated Google.")
    return se


def get_outlook_events(outlook_calendar):
    # get all events from an outlook calendar
    start_time = time.time()

    start = dt.datetime.today() - dt.timedelta(days=config.previous_days)
    end = dt.datetime.today() + dt.timedelta(days=config.future_days)
    events = outlook_calendar.get_events_in_range(start, end)

    elapsed_time = time.time() - start_time
    print("Retrieved {} events from Outlook in {:.1f} secs.".format(len(events), elapsed_time))
    return events


def clean_subject(subject):
    # remove prefix clutter from an outlook event subject
    remove = ["Fwd: ", "Invitation: ", "Updated invitation: ", "Updated invitation with note: "]
    for s in remove:
        subject = subject.replace(s, "")
    return subject


def clean_body(body):
    # strip out html and excess line returns from outlook event body
    text = BeautifulSoup(body, "html.parser").get_text()
    return text.replace("\n", " ").replace("\r", "\n")


def build_gcal_event(event):
    # construct a google calendar event from an outlook event

    e = {
        "summary": clean_subject(event.Subject),
        "location": event.Location,
        "description": clean_body(event.Body),
    }

    if event.AllDayEvent:
        # all day events just get a start/end date
        # use UTC start date to get correct day
        date = str(event.start.astimezone(pytz.utc).date())
        start_end = {"start": {"date": date}, "end": {"date": date}}
    else:
        # normal events have start/end datetime/timezone
        start = dateutil.parser.parse(str(event.start)) - dt.timedelta(hours=system_tz_offset)
        end = dateutil.parser.parse(str(event.end)) - dt.timedelta(hours=system_tz_offset)
        start_end = {
            "start": {
                "dateTime": str(start).replace(" ", "T"),
                #"timeZone": str(event.start.tzinfo),
                #"timeZone": str(config.system_default_tz),
            },
            "end": {
                "dateTime": str(end).replace(" ", "T"),
                #"timeZone": str(event.end.tzinfo),
                #"timeZone": str(config.system_default_tz),
            },
        }

    e.update(start_end)
    return e


def delete_google_events(se):
    # delete all events from google calendar
    start_time = time.time()
    gcid = config.google_calendar_id
    mr = 2500

    # retrieve a list of all events
    result = se.list(calendarId=gcid, maxResults=mr).execute()
    gcal_events = result.get("items", [])

    # if nextPageToken exists, we need to paginate: sometimes a few items are
    # spread across several pages of results for whatever reason
    i = 1
    while "nextPageToken" in result:
        npt = result["nextPageToken"]
        result = se.list(calendarId=gcid, maxResults=mr, pageToken=npt).execute()
        gcal_events.extend(result.get("items", []))
        i += 1

    print("Retrieved {} events across {} pages from Google.".format(len(gcal_events), i))

    # delete each event retrieved
    for gcal_event in gcal_events:
        request = se.delete(calendarId=config.google_calendar_id, eventId=gcal_event["id"])
        try:
            result = request.execute()
        except:# HttpError:
            # usually already been deleted
            try:
                print(f'Failed to delete {gcal_event["summary"]} on {gcal_event["start"]}, usually because its already been deleted.')
            except:
                print('Failed to delete something, but no idea why; summary or start time are missing.')
            pass
        assert result == ""
        time.sleep(config.pause)

    elapsed_time = time.time() - start_time
    print("Deleted {} events from Google in {:.1f} secs.".format(len(gcal_events), elapsed_time))


def add_google_events(se, events):
    # add all events to google calendar
    start_time = time.time()

    for event in events:
        e = build_gcal_event(event)
        result = se.insert(calendarId=config.google_calendar_id, body=e).execute()
        assert isinstance(result, dict)
        time.sleep(config.pause)

    elapsed_time = time.time() - start_time
    print("Added {} events to Google in {:.1f} secs.".format(len(events), elapsed_time))


def get_event_timestamps(outlook_events):
    # ids and timestamps of new events retrieved during current run
    ts = {}
    for e in outlook_events:
        ts[e.EntryID] = {
            "created_ts": int(e.CreationTime.timestamp()),
            "modified_ts": int(e.LastModificationTime.timestamp()),
        }
    return ts


def check_ts_match(new_events):
    # compare old event ids/timestamps to new ones retrieved during current run

    try:
        # load the old events' ids/timestamps saved to disk during previous run
        with open(config.events_ts_json_path, "r") as f:
            old_events = json.load(f)

        # make sure all ids and timestamps match between old and new
        assert new_events.keys() == old_events.keys()
        for k, new_event in new_events.items():
            old_event = old_events[k]
            assert new_event["created_ts"] == old_event["created_ts"]
            assert new_event["modified_ts"] == old_event["modified_ts"]

    except Exception:
        # if json file doesn't exist or if any id or timestamp is different
        print("Changes found.")
        return False

    return True


current_time = "{:%Y-%m-%d %H:%M:%S}".format(dt.datetime.now())
print("Started at {}.".format(current_time))
start_time = time.time()

# get local timezone offset in hours, e.g. EDT == -4
system_tz_offset = (time.timezone if (time.localtime().tm_isdst == 0) else time.altzone) / 60 / 60 * -1

# authenticate outlook and google credentials
outlook_calendar = outlook.outlookCal()
se = authenticate_google()

# get all events from outlook
outlook_events = get_outlook_events(outlook_calendar)
outlook_events_ts = get_event_timestamps(outlook_events)

# check if all the current event ids/timestamps match the previous run
# only update google calendar if they don't all match (means there are changes)
if config.force or not check_ts_match(outlook_events_ts):
    # delete all existing google events then add all outlook events
    delete_google_events(se)
    add_google_events(se, outlook_events)

    # save event ids/timestamps json to disk for the next run
    with open(config.events_ts_json_path, "w") as f:
        json.dump(outlook_events_ts, f)
else:
    print("No changes found.")

# all done
elapsed_time = time.time() - start_time
print("Finished in {:.1f} secs.\n".format(elapsed_time))

