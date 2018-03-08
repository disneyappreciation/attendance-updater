from __future__ import print_function

import os
import re

import chalk  # pip install pychalk
import httplib2
from apiclient import discovery
from dotenv import load_dotenv, find_dotenv  # pip install python-dotenv
from oauth2client import client  # pip install google-api-python-client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse

    flags = argparse.ArgumentParser(parents = [tools.argparser]).parse_args()
except ImportError:
    flags = None

load_dotenv(find_dotenv())

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Attendance Update Script'

spreadsheet_id = os.environ.get('SPREADSHEET_ID')
event_column = os.environ.get('EVENT_COLUMN')
sheet_name = os.environ.get('SHEET_NAME')
first_row = os.environ.get('FIRST_ROW')
last_row = os.environ.get('LAST_ROW')
first_column = os.environ.get('FIRST_COLUMN')
last_column = os.environ.get('LAST_COLUMN')


def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_row_number_of_person(spreadsheet, person):
    for i in range(len(spreadsheet)):
        item = spreadsheet[i]
        if len(item) == 0:
            continue

        if item[0].lower() == person[0].lower() and item[1].lower() == person[1].lower():
            return i
    return -1


def add_x_at(service, row, column):
    sheet_range = sheet_name + '!' + column + str(row + 3)
    value_input_option = 'RAW'
    value_range_body = {'values': [['x']]}
    request = service.spreadsheets().values().update(spreadsheetId = spreadsheet_id, range = sheet_range,
                                                     valueInputOption = value_input_option,
                                                     body = value_range_body)
    request.execute()


def insert_new_person(service, person):
    person_info = person[0:3]
    person_info.extend(['x' if n == (ord(event_column) - ord(first_column) - 3) else ''
                        for n in range(ord(last_column) - ord(first_column) - 3)])

    range_name = sheet_name + '!' + first_column + first_row + ':' + last_column + last_row
    value_input_option = 'RAW'
    value_range_body = {'values': [person_info]}
    request = service.spreadsheets().values().append(spreadsheetId = spreadsheet_id, range = range_name,
                                                     valueInputOption = value_input_option, body = value_range_body)
    request.execute()


def parse_file():
    csv = open(os.environ.get("FILE"), 'r')
    lines = csv.readlines()
    csv.close()
    new_records = [[strip_non_ascii(thing) for thing in line.split(',')] for line in lines]
    return new_records


def strip_non_ascii(s):
    r = re.sub(r'[^A-Za-z\'.@]', '', s)
    return r if isinstance(r, str) else r.decode('utf-8-sig')


def print_summary(already_accounted_for, not_in_spreadsheet, updated):
    chalk.red('People added to the spreadsheet:', opts = ('bold', 'underscore'))
    for person in not_in_spreadsheet:
        chalk.red('    - %s %s' % (person[0], person[1]))
    print()
    chalk.yellow('People already accounted for:', opts = ('bold', 'underscore'))
    for person in already_accounted_for:
        chalk.yellow('    - %s %s' % (person[0], person[1]))
    print()
    chalk.green('People updated:', opts = ('bold', 'underscore'))
    for person in updated:
        chalk.green('    - %s %s' % (person[0], person[1]))
    print()


def get_current_attendance(service):
    range_name = sheet_name + '!' + first_column + first_row + ':' + last_column + last_row
    result = service.spreadsheets().values().get(spreadsheetId = spreadsheet_id, range = range_name).execute()
    current_attendance = result.get('values', [])
    return current_attendance


def get_service():
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discovery_url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build('sheets', 'v4', http = http, discoveryServiceUrl = discovery_url)
    return service


def main():
    service = get_service()
    current_attendance = get_current_attendance(service)

    new_records = parse_file()
    not_in_spreadsheet = []
    already_accounted_for = []
    updated = []

    for record in new_records:
        row_num = get_row_number_of_person(current_attendance, record)
        column_num = ord(event_column) - ord('A')

        if row_num == -1:
            insert_new_person(service, record)
            not_in_spreadsheet.append(record)
        elif len(current_attendance[row_num]) > column_num and current_attendance[row_num][column_num] == 'x':
            already_accounted_for.append(record)
        else:
            add_x_at(service, row_num, event_column)
            updated.append(record)

    print_summary(already_accounted_for, not_in_spreadsheet, updated)


if __name__ == '__main__':
    main()
