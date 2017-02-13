"""
Reads through the computing bleep rota Excel file and
adds your slots to your Outlook calendar

"""

from openpyxl import load_workbook
import win32com.client
import datetime
import pytz
import config as cfg
import dateutil


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


def get_existing_appointments(outlook_object):
    """
    Function to read through all existing Outlook appts and filter
    out computing bleep appts
    :param outlook_object:
    :return: Returns a list of all existing Computing bleep appts
             to ensure that no duplicates are created
    """
    namespace = outlook_object.GetNamespace("MAPI")
    outlook_appointments = namespace.GetDefaultFolder(9).Items
    existing_appt_list = []
    for appt in outlook_appointments:
        data_dict = {'Start': [], 'Subject': []}
        data_dict['Start'] = dateutil.parser.parse(str(appt.Start)).replace(tzinfo=pytz.UTC)
        data_dict['Subject'] = appt.Subject
        existing_appt_list.append(data_dict)
    return [element for element in existing_appt_list if 'Computing bleep' in element['Subject']]


def convert_dates_to_appointments(input_data, user_initials):
    """
    Function to take list of dictionaries and create Outlook appointments
    where your initials are in the rota slot
    :param input_data: From above function
    :param user_initials: Your initials e.g. 'LG'
    :return: None
    """
    count = 0
    existing_appts = get_existing_appointments(outlook)
    for entry in input_data:
        try:
            if user_initials in entry['AM']:                   # if your initials are present in AM
                if entry['Date'] > datetime.datetime.now():    # and calendar entry is after today (don't care
                    if user_initials in entry['AM']:
                        appointment = outlook.CreateItem(1)
                        appointment.Start = entry['Date'].replace(tzinfo=pytz.UTC) +\
                                            datetime.timedelta(hours=8)         # Starts at 8 AM
                        if not [item for item in existing_appts if item['Start'] == appointment.Start]:
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
                        if not [item for item in existing_appts if item['Start'] == appointment.Start]:
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
    file_path = cfg.file_path
    sheet_label = cfg.sheet_label
    excel_data = parse_workbook(file_path, sheet_label)
    user = cfg.user
    convert_dates_to_appointments(excel_data, user)
