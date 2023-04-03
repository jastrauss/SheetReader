"""
    An example of using SheetReader to read the names in a
    Google Spreadsheet and write back a greeting.
"""

from SheetReader import SheetReader

reader = SheetReader(
    workbook_id="",  # The URL of the Google Sheet
    sheet_name="Sheet1",  # The name of the sheet tab
    header_range=("A1", "C1"),  # Range with the headers, one row
    data_range=("A2", "C3"),  # Range of the data
)

print(reader.headers)

for row in reader:
    # Read the name
    name = row["Name"]
    print("Hello", name)

    # Write a greeting
    row["Greeting"] = f"Hello {name}!"

    # Write the reversed name
    row["Reversed Name"] = name[::-1].lower()
