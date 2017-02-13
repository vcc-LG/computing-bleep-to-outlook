"""
Reads through the computing bleep rota Excel file and
adds your slots to your Outlook calendar

"""

from openpyxl import load_workbook
import win32com.client
import datetime
import pytz
import config as cfg


def parse_workbook(path_to_file, sheet_name):
    """
    Function to read in an xlsx file
    :param path_to_file: The full path to the Excel file
    :param sheet_name: The name of the sheet to read in
    :return: all_dates: A list of dictionaries of Dates and Initials
                        for AM and PM slots
    """
    wb = load_workbook(path_to_file, read_only=True)
    ws = wb[sheet_name]
    all_dates = []
    for row in ws.rows:
        data_dict = {'Date': [], 'AM': [], 'PM': []}
        try:
            if isinstance(row[0].value, datetime.datetime):
                data_dict['Date'] = row[0].value
                data_dict['AM'] = row[2].value
                data_dict['PM'] = row[3].value
                all_dates.append(data_dict)
        except:
            pass
    return all_dates


def remove_existing_appointments(outlook_object):
    """
    Function to read through all existing Outlook appts and remove
    existing computing bleep appts to avoid dupes and allow for
    changes to the schedule
    :param outlook_object:
    :return: None
    """
    namespace = outlook_object.GetNamespace("MAPI")
    outlook_appointments = namespace.GetDefaultFolder(9).Items
    for appt in outlook_appointments:
        if 'Computing bleep' in appt.Subject:
            appt.Delete()
    return None


def convert_dates_to_appointments(input_data, user_initials):
    """
    Function to take list of dictionaries and create Outlook appointments
    where your initials are in the rota slot
    :param input_data: From above function
    :param user_initials: Your initials e.g. 'LG'
    :return: None
    """
    count = 0
    for entry in input_data:
        try:
            if user_initials in entry['AM']:                   # if your initials are present in AM
                if entry['Date'] > datetime.datetime.now():    # and calendar entry is after today (don't care
                    if user_initials in entry['AM']:
                        appointment = outlook.CreateItem(1)
                        appointment.Start = entry['Date'].replace(tzinfo=pytz.UTC) +\
                                            datetime.timedelta(hours=8)         # Starts at 8 AM
                        appointment.Subject = 'Computing bleep AM'
                        appointment.Duration = 5 * 60                      # Lasts 5 hours
                        appointment.ReminderSet = True
                        appointment.ReminderMinutesBeforeStart = 15
                        appointment.Save()
                        count += 1
                    if user_initials in entry['PM']:
                        appointment = outlook.CreateItem(1)
                        appointment.Start = entry['Date'].replace(tzinfo=pytz.UTC) + \
                                            datetime.timedelta(hours=13)      # Starts at 1PM
                        appointment.Subject = 'Computing bleep PM'
                        appointment.Duration = 5 * 60
                        appointment.ReminderSet = True
                        appointment.ReminderMinutesBeforeStart = 15
                        appointment.Save()
                        count += 1
        except TypeError:
            pass
    print("Added {0} calendar entries".format(count))
    return None

if __name__ == '__main__':
    outlook = win32com.client.Dispatch("Outlook.Application")
    remove_existing_appointments(outlook)
    remove_existing_appointments(outlook)   #has to run twice because of the magic
    file_path = cfg.file_path
    sheet_label = cfg.sheet_label
    excel_data = parse_workbook(file_path, sheet_label)
    user = cfg.user
    convert_dates_to_appointments(excel_data, user)
