import win32com.client
from datetime import datetime, timedelta

class outlookCal(object):

    def __init__(self):
        # Initialize the Outlook application
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        # Get the calendar folder
        self.calendar_folder = namespace.GetDefaultFolder(9)  # 9 corresponds to the Calendar
        return

    def get_events_in_range(self, dt_start, dt_end):
        events = []
        # Filter appointments between the start and end time
        appointments = self.calendar_folder.Items
        appointments.IncludeRecurrences = True
        appointments.Sort("[Start]")

        # Restrict to only today's appointments
        restriction = "[Start] >= '{0}' AND [Start] < '{1}'".format(
            dt_start.strftime("%m/%d/%Y %H:%M %p"),
            dt_end.strftime("%m/%d/%Y %H:%M %p")
        )
        inrange_appointments = appointments.Restrict(restriction)

        # List today's appointments and meetings
        for appointment in inrange_appointments:
            # Check if the event is recurring
            if appointment.IsRecurring:
                try:
                    recurrence_pattern = appointment.GetRecurrencePattern()
                    occurrence = recurrence_pattern.GetOccurrence(dt_start)

                    if occurrence:
                        appointment = occurrence  # Get the specific occurrence details
                except Exception as e:
                    #print(f"Warning: Could not retrieve occurrence for {appointment.Subject} due to: {str(e)}")
                    pass

            events.append(appointment)

        return events

    def get_events_today(self):
        today = datetime.now()
        return self.get_events_in_range(today, today)

    def get_teams_link(self, appointment):
        try:
            link = re.search('(?<=Join the meeting now).*', appointment.Body)[0].strip()
            return link
        except:
            return ""

    def todo__print__(self, appointment):
        print(f"{meeting_type}: {appointment.Subject}")
        print(f"Start: {appointment.Start}")
        print(f"End: {appointment.End}")
        print(f"Teams Link: {teams_link if teams_link else 'No Teams link found'}")
        print(f"EntryID: {appointment.EntryID}")
        print("\n")
        return


