# Copyright 2018 Ethan Genser

# Importing packages, modules, and functions.
import xlwt
from os.path import basename
from datetime import datetime
import pandas as pd

def main():
    sheets = 0
    restart = True
    timestamp = datetime.now()
    print_copyright()
    book = xlwt.Workbook(encoding='utf-8')

    while restart:
        # Opens data file.
        path = input('\nWhat file what you like to scan?\n>>> ')
        data_file = open(path, 'r')
        raw_data = data_file.readlines()

        # Reads and collects data from the file.
        print('\nCollecting data...\n')
        data = get_data(raw_data)

        # Opens a new worksheet. 
        sheets += 1
        sheet = book.add_sheet('Sheet' + str(sheets))
        sheet = format_sheet(sheet, timestamp, raw_data)

        # Records collected data in spreadsheet.
        sheet = record_data(sheet, data)

        # Asks if the user wants to scan another file.
        restart = False
        restart_raw = ''
        while restart_raw.casefold() != 'y' and restart_raw.casefold() != 'n':
            restart_raw = input('\nYou you like to scan another file (Y or N)?\n>>>')
            if restart_raw.casefold() == 'y':
                restart = True
            elif restart_raw.casefold() == 'n':
                restart = False

    # Saves finished workbook
    book.save('C:\\Users\\Ethans Laptop\\Desktop\\CDC_' + str(timestamp.month) + '-' + str(timestamp.day) + '-' + str(timestamp.year) + '.xls')

def get_data(raw_data:list)->list:

    def get_causes(raw_data:list)->list:
        causes = list()
        cause = ''
        rec = False
        line_number = 0

        # Iterates through each line of the file.
        for line in raw_data:
            line_number += 1
            cause = ''
            # Skips the first line.
            if line_number < 2:
                continue
            # Stops reading when the dividor is reached.
            if line == '\"---\"\n':
                break

            # Iterates through each character in the line.
            for char in line:
                # Stops recording characters when the stop symbol is reached.
                if char == '(':
                    rec = False
                # Records characters.
                if rec:
                    cause = cause + char
                # Starts recording characters when the start symbol is reached.
                if char == '#':
                    rec = True

            # Formats and appends the cause to the list of causes.
            cause = cause.strip()
            causes.append(cause)

        return causes
    def get_deaths(raw_data:list)->list:
        deaths = list()
        total = ''
        tabs = 0
        rec = False
        line_number = 0

         # Iterates through each line of the file.
        for line in raw_data:
            line_number += 1
            total = ''
            tabs = 0
            # Skips the first line.
            if line_number < 2:
                continue
            # Stops reading when the dividor is reached.
            if line == '\"---\"\n':
                break

            # Iterates through each character in the line.
            for char in line:
                # Counts the number of tabs since the beginning of the line.
                if char == '\t':
                    tabs += 1
                # Starts recording characters when the second tab is reached.
                if tabs > 2:
                    rec = True
                # Stops recording characters when the third tab is reached.
                if tabs > 3:
                    rec = False
                # Records characters.
                if rec:
                    total = total + char

            # Formats and appends the total to the list of death totals.
            total = int(total.strip())
            deaths.append(total)

        return deaths

    data = list()
    data.append(get_causes(raw_data))
    data.append(get_deaths(raw_data))
    return data

def format_sheet(sheet:xlwt.Worksheet, timestamp:datetime, raw_data:list)->xlwt.Worksheet:

    def get_time(raw_data:list)->str:
            time = ''
            rec = False

            # Iterates through each line of the file.
            for line in raw_data:

                if 'Year/Month' in line:
                    # Iterates through each character in the line.
                    for char in line:
                        # Stops recording characters when the stop symbol is reached.
                        if char == '\"':
                            rec = False
                        # Records characters.
                        if rec:
                            time = time + char
                        # Starts recording characters when the start symbol is reached.
                        if char == ':':
                            rec = True

                    # Formats the string
                    time = time.strip()
                    break

            return time

    sheet.write(0,3,'Spreadsheet generated by ' + basename(__file__) + " on " + str(timestamp.month) + '/' + str(timestamp.day) + '/' + str(timestamp.year) + ' @ ' + str(timestamp.hour) + ':' + str(timestamp.minute) + ':' + str(timestamp.second))
    sheet.write(2,3,'Dataset provided by the Center for Disease Control and Prevention (https://wonder.cdc.gov/)')
    sheet.write(0,0,'Cause of Death')
    sheet.write(0,1,'Total Fatalities (' + get_time(raw_data) + ')')
    return sheet

def record_data(sheet:xlwt.Worksheet, data:list)->xlwt.Worksheet:
    x = 0
    y = 1
    for column in data:
        y = 1
        for row in column:
            sheet.write(y,x,row)
            y += 1
        x += 1
    return sheet

def print_copyright():
    print('**********************************************')
    print('* ' + basename(__file__) + ' (c) Ethan Genser 2018 *')
    print('**********************************************')

if __name__=='__main__':main()