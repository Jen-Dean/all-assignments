#Import Modules

import os
import csv
from datetime import datetime

#Declare Variables
Em_ID = []

FirstName = []
LastName = []

Month = []
Day = []
Year = []
BirthDate = []

SSNumber_1 = []
SSNumber_2 = []
SSNumber_3 = []
SSNumber_All = []

State = []

SplitName = 0
Birthday = 0
Raw_SSN = 0
dob = 0

#US State Abbreviation Dictionary
us_state_abbrev = {
    'Alabama': 'AL',
    'Alaska': 'AK',
    'Arizona': 'AZ',
    'Arkansas': 'AR',
    'California': 'CA',
    'Colorado': 'CO',
    'Connecticut': 'CT',
    'Delaware': 'DE',
    'Florida': 'FL',
    'Georgia': 'GA',
    'Hawaii': 'HI',
    'Idaho': 'ID',
    'Illinois': 'IL',
    'Indiana': 'IN',
    'Iowa': 'IA',
    'Kansas': 'KS',
    'Kentucky': 'KY',
    'Louisiana': 'LA',
    'Maine': 'ME',
    'Maryland': 'MD',
    'Massachusetts': 'MA',
    'Michigan': 'MI',
    'Minnesota': 'MN',
    'Mississippi': 'MS',
    'Missouri': 'MO',
    'Montana': 'MT',
    'Nebraska': 'NE',
    'Nevada': 'NV',
    'New Hampshire': 'NH',
    'New Jersey': 'NJ',
    'New Mexico': 'NM',
    'New York': 'NY',
    'North Carolina': 'NC',
    'North Dakota': 'ND',
    'Ohio': 'OH',
    'Oklahoma': 'OK',
    'Oregon': 'OR',
    'Pennsylvania': 'PA',
    'Rhode Island': 'RI',
    'South Carolina': 'SC',
    'South Dakota': 'SD',
    'Tennessee': 'TN',
    'Texas': 'TX',
    'Utah': 'UT',
    'Vermont': 'VT',
    'Virginia': 'VA',
    'Washington': 'WA',
    'West Virginia': 'WV',
    'Wisconsin': 'WI',
    'Wyoming': 'WY',
}

#join the file information to read from
data_csv = os.path.join("PyBoss/Resources/employee_data.csv")

with open(data_csv) as csvfile:
    csv_reader = csv.reader(csvfile, delimiter=",")
    csv_header = next(csv_reader)

#Create a Loop to place all Variables into the proper lists
    for row in csv_reader:
        Em_ID.append(row[0])

        #The Name column should be split into separate First Name and Last Name columns.
        SplitName = row[1].split(" ")
        FirstName.append(SplitName[0])
        LastName.append(SplitName[1])

        ############Birthday = row[2].split("-")
        ############Year.append(Birthday[0])
        ############Month.append(Birthday[1])
        ############Day.append(Birthday[2])

        #The DOB data should be re-written into MM/DD/YYYY format.
        date_time_object = datetime.strptime(row[2], '%Y-%m-%d')
        dob = date_time_object.strftime('%m/%d/%Y')
        BirthDate.append(dob)

        #The SSN data should be re-written such that the first five numbers are hidden from view.
        Raw_SSN = row[3].split("-")
        Last5 = Raw_SSN[2]
        SSNumber_All.append(f"***-**-{Last5}")

        #The State data should be re-written as simple two-letter abbreviations.
        State_long = row[4]
        State.append(us_state_abbrev[State_long])

#Join back all data using a zip
Final_File = zip(Em_ID,FirstName,LastName,BirthDate,SSNumber_All,State)

#Merge the List and Export into a csv

# Specify the file to write to
output_path = os.path.join("/Users/jenniferdean/Desktop/DATA_SCIENCE_BOOTCAMP/GitHub/HOMEWORK/python-challenge/PyBoss/output/new_employee_data.csv")

# Open the file using "write" mode. Specify the variable to hold the contents
with open(output_path, 'w', newline = "") as csvfile:

    # Initialize csv.writer
    csvwriter = csv.writer(csvfile, delimiter=',')

    # Write the first row (column headers)
    csvwriter.writerow(['Emp ID', 'First Name', 'Last Name', 'DOB', 'SSN', 'State'])

    csvwriter.writerows(Final_File)