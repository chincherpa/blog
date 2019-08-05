Title: [Python] create all day Outlook-event
Date: 2019-06-03 05:57
Author: Lulef
Category: Sammlung
Slug: python-create-all-day-outlook-event
Status: published
```
import win32com.client
def add_all_day_event(start, subject, loc, duration=1, reminder=False, reminder_days=1):
    """
    :param start:    date
    :param subject:  subject
    :param loc:      location
    :param duration: duration in days
    """
    appointment = win32com.client.Dispatch('Outlook.Application').CreateItem(1)
    appointment.AllDayEvent = True
    appointment.Start = start
    appointment.Subject = subject
    appointment.Location = loc
    appointment.Duration = 1440 * duration
    appointment.ReminderSet = reminder
    appointment.ReminderMinutesBeforeStart = 1440 * reminder_days
    appointment.Save()
def folder_tree(folders, indent=0):
    prefix = ' ' * (indent * 2)
    i = 0
    for folder in folders:
        print(f'{prefix}{i}. {folder.Name} ({folder.DefaultItemType})')
        folder_tree(folder.Folders, indent + 1)
        i += 1
if __name__ == '__main__':
    if 1:
        start = '2019-06-05'
        subject = 'Conference'
        loc = 'Venice'
        duration = 2
        add_all_day_event(start, subject, loc, duration)
    else:
        namespace = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        folder_tree(namespace.Folders)
```
