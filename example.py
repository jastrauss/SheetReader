"""
    An example of using SheetReader to reverse all the names in a
    Google Spreadsheet, and calculate ages from birthdays.
"""

import datetime
from dateutil.parser import parse
from SheetReader import SheetReader
:
""" Returns age in years """
def get_age_from_birthday(birthday):
    today = datetime.date.today()
    years = today.year - birthday.year
    if today.month < birthday.month or (today.month == birthday.month and today.day < birthday.day):
        years -= 1
    return years

workbook_id = '' # The ID of the workbook, https://docs.google.com/spreadsheets/d/THIS_STRING_IS_THE_WORKBOOK_ID/edit
sheet_name = 'Sheet1' # The name of the specfic sheet
header_range = ('A1','D1') # Range with the headers, one row
data_range = ('A2','D10') # Range of the data

reader = SheetReader(workbook_id=workbook_id, sheet_name=sheet_name, header_range=header_range, data_range=data_range)
print reader.headers

for row in reader:

    name = row["name"]
    print "Hello", name

    # Reverse the name
    row["Reversed Name"] = name[::-1]

    # Calculate their age
    parsed_birthday = parse(row["Birthday"])
    row["Current Age"] = get_age_from_birthday(parsed_birthday)

