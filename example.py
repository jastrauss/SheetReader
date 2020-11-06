"""
    An example of using SheetReader to read the names in a
    Google Spreadsheet and write back a greeting.
"""

from SheetReader import SheetReader

workbook_id = '' # The ID of the workbook, https://docs.google.com/spreadsheets/d/THIS_STRING_IS_THE_WORKBOOK_ID/edit
sheet_name = 'Sheet1' # The name of the specfic sheet
# See the example spreadsheet in example_spreadsheet.png
header_range = ('A1','C1') # Range with the headers, one row
data_range = ('A2','C3') # Range of the data

reader = SheetReader(
    workbook_id=workbook_id,
    sheet_name=sheet_name,
    header_range=header_range,
    data_range=data_range)

print(reader.headers)

for row in reader:

    # Read the name
    name = row["Name"]
    print("Hello", name)

    # Write a greeting
    row["Greeting"] = f"Hello {name}!"

    # Write the reversed name
    row["Reversed Name"] = name[::-1].lower()
