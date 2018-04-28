# Copyright 2018 Ethan P. Genser
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#    http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# Importing packages, modules, and functions.
import xlsxwriter
from os.path import basename

# Constants
SAVE_PATH = 'C:/Users/Ethans Laptop/Desktop/CDC_Data.xlsx'
YEARS = range(2008,2017)
DATA_BASE_NAME = 'Cause_of_Death_'

# The main entry point for the program.
def main():
    workbook = xlsxwriter.Workbook(SAVE_PATH)
    raw_data = list()
    causes = list()
    deaths = list()

    # Displays copyright info.
    print_copyright()

    # Opens and reads each data file.
    for year in YEARS:
        raw_data.append(open(str(DATA_BASE_NAME) + str(year) + '.txt', 'r').readlines())

    # Collects data from each file.
    print('\nCollecting data...\n')
    for i in range(0, len(YEARS)):
        causes.append(get_causes(raw_data[i]))
        deaths.append(get_deaths(raw_data[i]))

    # Records data in the spreadsheet.
    worksheet = workbook.add_worksheet()
    worksheet = record_data(workbook, worksheet, causes, deaths)

    # Saves finished workbook.
    workbook.close()

# Returns a nested list of the causes of death detailed in each data file.
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

# Returns a nested list of the total number of fatalities for each cause.
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

# Writes the collected information into an xlsx file.
def record_data(workbook:xlsxwriter.workbook, worksheet:xlsxwriter.worksheet, causes:list, deaths:list)->xlsxwriter.worksheet:

    # Creates a histogram of the collected data.
    def create_histogram():

        # Sets the graph's metadata.
        chart = workbook.add_chart({'type': 'column'})
        chart.set_y_axis({'name': 'Fatalities'})
        chart.set_x_axis({'name': 'Cause of Death'})
        chart.set_title({'name': '15 Leading Causes of Death in America'})

        # Adds a new data series to the chart for every year.
        for year in YEARS:

            # Finds the top coordinate of the applicable data.
            top = ((year - YEARS[0]) * (len(causes[0]) + 1)) + 1

            # Finds the bottom coordinate of the applicable data.
            bottom = ((year - YEARS[0]) * (len(causes[0]) + 1)) + len(causes[0])

            # Creates the new series.
            chart.add_series({
            'categories': '=Sheet1!$A$' +  str(top) + ':$A$' + str(bottom), # Causes
            'values': '=Sheet1!$B$' +  str(top) + ':$B$' + str(bottom),     # Deaths
            'name': str(year),                                              # Year
            })
        worksheet.insert_chart('D1', chart)

    # Writes the causes of death for each year into the spreadsheet.
    for i in range(0, len(causes)):
        worksheet.write_column(*[i * (len(causes[i]) + 1), 0], data=causes[i])

    # Writes the death total for each cause into the spreadsheet.
    for i in range(0, len(deaths)):
        worksheet.write_column(*[i * (len(deaths[i]) + 1), 1], data=deaths[i])

    # Visualizes data as a histogram
    create_histogram()

# Prints copyright info to the comandline terminal.
def print_copyright():
    print('**********************************************')
    print('* ' + basename(__file__) + ' (c) Ethan Genser 2018 *')
    print('**********************************************')

if __name__=='__main__':main()
