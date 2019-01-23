import os
import re

import chalk
import httplib2
from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import argparse

flags = argparse.ArgumentParser(parents = [tools.argparser]).parse_args()

SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Attendance Update Script'

spreadsheet_id = os.environ.get('SPREADSHEET_ID')
event_column = os.environ.get('EVENT_COLUMN')
sheet_name = os.environ.get('SHEET_NAME')

spreadsheet = []
service = None


###########################
# Update Global Variables #
###########################


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
        credentials = tools.run_flow(flow, store, flags)
        print('Storing credentials to ' + credential_path)
    return credentials


def get_current_attendance():
    global spreadsheet
    if spreadsheet:
        return spreadsheet

    range_name = sheet_name + '!A3:' + event_column
    result = get_service().spreadsheets().values() \
        .get(spreadsheetId = spreadsheet_id, range = range_name).execute()
    current_attendance = result.get('values', [])

    spreadsheet = current_attendance
    return current_attendance


def get_service():
    global service
    if service:
        return service

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discovery_url = 'https://sheets.googleapis.com/$discovery/rest?version=v4'
    service = discovery.build('sheets', 'v4', http = http, discoveryServiceUrl = discovery_url)
    return service


#####################
# Parsing CSV Files #
#####################


def parse_file():
    csv = open(os.environ.get("FILE"), 'r')
    lines = csv.readlines()
    csv.close()
    return [[strip_non_ascii(thing) for thing in line.split(',')] for line in lines]


def strip_non_ascii(s):
    r = re.sub(r'[^A-Za-z\'.@]', '', s)
    return r if isinstance(r, str) else r.decode('utf-8-sig')


##############################
# Spreadsheet Helper Methods #
##############################

def get_row_number_of_person(person):
    first_name = 0
    last_name = 1

    for i in range(len(spreadsheet)):
        item = spreadsheet[i]
        if len(item) == 0:
            continue

        if item[first_name].lower() == person[first_name].lower() and \
           item[last_name].lower() == person[last_name].lower():
            return i
    return -1


def update_spreadsheet():
    range_name = sheet_name + '!A3:' + event_column + str(len(spreadsheet) + 4)
    value_input_option = 'RAW'
    value_range_body = {'values': spreadsheet}
    request = service.spreadsheets().values().append(spreadsheetId = spreadsheet_id, range = range_name,
                                                     valueInputOption = value_input_option, body = value_range_body)
    request.execute()


########
# Main #
########


def print_summary(already_accounted_for, not_in_spreadsheet, updated):
    print chalk.red('People added to the spreadsheet:', bold = True, underline = True)
    for person in not_in_spreadsheet:
        print chalk.red('    - %s %s' % (person[0], person[1]))
    print
    print chalk.yellow('People already accounted for:', bold = True, underline = True)
    for person in already_accounted_for:
        print chalk.yellow('    - %s %s' % (person[0], person[1]))
    print
    print chalk.green('People updated:', bold = True, underline = True)
    for person in updated:
        print chalk.green('    - %s %s' % (person[0], person[1]))
    print


def main():
    get_current_attendance()

    new_records = parse_file()
    not_in_spreadsheet = []
    already_accounted_for = []
    updated = []

    for record in new_records:
        row_num = get_row_number_of_person(record)
        column_num = ord(event_column) - ord('A')

        if row_num == -1:
            row = record[0:3]
            row.extend(['x' if n == (ord(event_column) - ord('A') - 3) else ''
                       for n in range(ord(event_column) - ord('A') - 2)])
            spreadsheet.append(row)
            not_in_spreadsheet.append(record)
        elif len(spreadsheet[row_num]) > column_num and spreadsheet[row_num][column_num] == 'x':
            already_accounted_for.append(record)
        else:
            spreadsheet[row_num][column_num] = 'x'
            updated.append(record)

    update_spreadsheet()
    print_summary(already_accounted_for, not_in_spreadsheet, updated)


if __name__ == '__main__':
    main()
