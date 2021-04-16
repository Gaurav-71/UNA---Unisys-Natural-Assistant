import win32com.client
import datetime
from collections import namedtuple

from O365 import Account, MSGraphProtocol





def parse_event(event):
    event_str = str(event)
    start_index = event_str.find('from') + 6
    start = event_str[start_index: start_index + 8]
    start_obj = datetime.datetime.strptime(start, '%H:%M:%S')
    end_index = event_str.find('to') + 4
    end = event_str[end_index: end_index + 8]
    end_obj = datetime.datetime.strptime(end, '%H:%M:%S')
    sub_index = event_str.find(':')
    brac_index = event_str.find('(')
    subject = event_str[sub_index+1:brac_index]

    date = event_str[event_str.find("on: ") + 4: event_str.find("on: ") + 14]

    new_event = {"date": date, "subject": subject, "from" : start_obj.time(), "to" : end_obj.time()}
    return(new_event)


'''q = calendar.new_query('start').greater_equal(datetime.datetime.today())
q.chain('and').on_attribute('end').less_equal(datetime.timedelta(days=1) + datetime.datetime.today())'''
def show_calendar(account, endDate):
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()

    q = calendar.new_query('start').greater_equal(datetime.datetime(datetime.date.today().year, datetime.date.today().month, datetime.date.today().day, 5, 30))
    q.chain('and').on_attribute('end').less_equal(endDate)
    events = calendar.get_events(query = q, include_recurring=False)
    
    for event in events:
        new_event = parse_event(event)
        print(new_event["subject"], new_event["from"], new_event["to"])

def add_event(account):
    schedule = account.schedule()
    calendar = schedule.get_default_calendar()
    new_event = calendar.new_event()
    new_event.subject = 'Adding new one'
    new_event.start = datetime.datetime(2021, 2, 28, 14, 30)
    new_event.end = datetime.datetime(2021, 2, 28, 15, 0)
    new_event.save()