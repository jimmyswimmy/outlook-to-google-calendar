# Microsoft Office 365 calendar -> Google calendar

## Overview

One-way sync from a Microsoft Office Outlook local calendar to a Google calendar, handling new, updated, and deleted events.

The script connects to the local instance of Outlook using [pywin32]|(https://pypi.org/project/pywin32) and connects to the Google API using its [Python client](https://developers.google.com/calendar/api/quickstart/python). Familiarize yourself with their documentation as you may need to enable APIs or create credentials per their instructions before you begin. See also the Google Calendar API [reference](https://developers.google.com/calendar/v3/reference/events).

## Setup

  - Create `config.py` (you can adapt [`config_sample.py`](config_sample.py)) to hold your personal configuration details, include your Microsoft `client_id` and `client_secret` and your Google calendar ID.
      - **Create a new Google calendar just for this application, or else your existing events will be deleted!**
      - **no seriously, it defaults to deleting every single event on your google calendar!**
      - ** SERIOUSLY read the last two lines **
      - Create Google credentials for this application (see overview section above) and save as `credentials/google_credentials.json`.
  - Run `pip install --upgrade -r requirements.txt` to install the [required Python dependencies](requirements.txt).
  - In the [credentials folder](credentials), run [`python quickstart.py`](credentials/quickstart.py) to create a Google API access token.
  - On your server, set up a cron job to run [`outlook_to_google.py`](outlook_to_google.py) (using [run.sh](run.sh)) every 15 minutes (or however often you need).
  - The script will check Outlook for calendar events and compare them to the calendar events it saved (in events_ts.json) during the previous run. **If they differ (in IDs or timestamps), it will delete all events on this Google calendar and then add all Microsoft calendar events to the Google calendar.**

## TODO
  - ~~recognize current system time zone and take that offset out of config.py~~
  - find teams links and put them back into the body, or make new teams links from the meeting id
  - make recurring meetings... recurring meetings
  - be smarter about tracking changes
  - only update changed stuff, rather than deleting everything and recreating
  - be smarter about time zones - if tz changes during lookahead period, adjust for it
