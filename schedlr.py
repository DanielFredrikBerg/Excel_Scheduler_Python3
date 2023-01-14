#!/usr/bin/python3

import datetime
import sys
import csv

import json

# TODO: Finns TIMEEDIT API https://developer.timeedit.com/#/objects?id=getobjects
# Kolla vad som skickas och vad som returneras i BurpSuite.


# TODO: testfile for this function
def get_week(date, date_format='ymd', delimiter='-'):
    match date_format:
        case 'ymd':
            year, month, day = [int(x) for x in date.split(delimiter)]
        case 'ydm':
            year, day, month = [int(x) for x in date.split(delimiter)]
        case 'dmy':
            day, month, year = [int(x) for x in date.split(delimiter)]
    week = datetime.date(year, month, day).strftime("%W")
    if week == '00':
        return '1'
    else: return week.lstrip('0')
    

def format_timeedit_categories(csv, start_line=3):
    with open(csv, 'r+') as csv_file:
        lines = csv_file.readlines()

    with open('try2.csv', 'w') as csv_file:
        csv_file.writelines(lines[start_line:])

        
def initialize_week_schedule(file_name, relevant_activities,
                             wanted_info = ['Kurs', 'Undervisningstyp', 'Information till student']):
    with open(file_name, 'r') as csv_file:
        week_schedule = {}
        csv_reader = csv.DictReader(csv_file)
        counter = 1
        for row in csv_reader:
            if row != '':
                if row['Undervisningstyp'] in relevant_activities[row['Kurs']]:
                    if row['Kurs'] == 'TDP007' and row['Undervisningstyp'] == 'Seminarium':
                        row['Undervisningstyp'] = 'Seminarium ' + str(counter)
                        counter += 1
                    
                    activity_date = row['Startdatum']                    
                    week = get_week(activity_date)
                    work_time = f"{row['Starttid']}-{row['Sluttid']}"

                    if week not in week_schedule:
                        week_schedule[week] = {}
                    if week in week_schedule \
                       and activity_date not in week_schedule[week]:
                        week_schedule[week][activity_date] = {}
                    if week in week_schedule and activity_date in week_schedule[week]:
                        week_schedule[week][activity_date][work_time] = [ row[info] for info in wanted_info ]
    return week_schedule

# TODO: CLI setup for wanted values?
relevant_activities = {
    'TDP007': ['Laboration', 'Seminarium', 'Dugga'],
    'TDP019': ['Handledning'],
}


# wanted_info = ['Kurs', 'Undervisningstyp', 'Information till student']

ws = initialize_week_schedule(sys.argv[1], relevant_activities)
print(json.dumps(ws, indent=2))

