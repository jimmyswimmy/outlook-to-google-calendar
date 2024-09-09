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

    def get_events_in_range(self, dt_start, dt_end, include_recurrences=True):
        events = []
        # Filter appointments between the start and end time
        appointments = self.calendar_folder.Items
        appointments.IncludeRecurrences = include_recurrences
        appointments.Sort("[Start]")

        # Restrict to only today's appointments
        restriction = "[Start] >= '{0}' AND [Start] < '{1}'".format(
            dt_start.strftime("%m/%d/%Y %H:%M %p"),
            dt_end.strftime("%m/%d/%Y %H:%M %p")
        )
        inrange_appointments = appointments.Restrict(restriction)
        return inrange_appointments

    def get_all_events_in_range(self, dt_start, dt_end):
        events = []
        inrange_appointments = self.get_events_in_range(dt_start, dt_end)
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

    def get_nonrecurring_events(self, dt_start, dt_end, nonrecurring=True, include_recurrences=True):
        events = []
        inrange_appointments = self.get_events_in_range(dt_start, dt_end, include_recurrences=include_recurrences)
        # List today's appointments and meetings
        for appointment in inrange_appointments:
            # Check if the event is recurring
            if nonrecurring:
                if not appointment.IsRecurring:
                    events.append(appointment)
            else:
                if appointment.IsRecurring:
                    events.append(appointment)
        return events

    def get_recurring_events(self, dt_start, dt_end):
        return self.get_nonrecurring_events(dt_start, dt_end, nonrecurring=False, include_recurrences=False)

    def get_events_today(self):
        today = datetime.now()
        return self.get_events_in_range(today, today)

    def get_teams_link(self, appointment):
        try:
            link = re.search('(?<=Join the meeting now).*', appointment.Body)[0].strip()
            return link
        except:
            return ""

    def parse_recurring_event(self, event):
        if event.IsRecurring():
            # get recurrence type
            # https://learn.microsoft.com/en-us/office/vba/api/outlook.recurrencepattern.recurrencetype
            recurrence = event.GetRecurrencePattern()
            recurrence_type_id = recurrence.RecurrenceType
            if recurrence_type_id == 0:
                #daily
                recurrence_type = 'daily'
            if recurrence_type_id == 1:
                #weekly
                recurrence_type = 'weekly'
                # https://learn.microsoft.com/en-us/office/vba/api/outlook.oldaysofweek
                # sunday = 1, monday = 2, tues = 4, weds = 8, thurs = 16, fri = 32, sat = 64
                dayMask = recurrence.DayOfWeekMask
            if recurrence_type_id == 2:
                #monthly
                recurrence_type = 'monthly'
                dayOfMonth = recurrence.DayOfMonth
            if recurrence_type_id == 3:
                #monthly nth day
                recurrence_type = 'monthly_nthday'
                dayMask = recurrence.DayOfWeekMask
                instance = recurrence.Instance # nth week of the month; values > 5 are errors

    def todo__print__(self, appointment):
        print(f"{meeting_type}: {appointment.Subject}")
        print(f"Start: {appointment.Start}")
        print(f"End: {appointment.End}")
        print(f"Teams Link: {teams_link if teams_link else 'No Teams link found'}")
        print(f"EntryID: {appointment.EntryID}")
        print("\n")
        return


