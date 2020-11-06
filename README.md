# SheetReader

Iterate through rows of Google Spreadsheet as if it were a Python CSV DictReader or DictWriter.

## Getting Started
See [example.py](example.py) for a full working example.

![Image of Spreadsheet with Chicken Wing Prices](/chicken_wing_spreadsheet.png)
```python
reader = SheetReader(
    workbook_id = 'workbook_id_123',
    sheet_name = 'WingPrices',
    header_range=('A1', 'C1'),
    data_range = ('A2','C10')
)

for row in reader:
    # Read row values
    num_wings = row['Chicken Wings']
    price = row['Price']

    # Write a value
    row['Price per Wing'] = float(price) / int(num_wings)
```

## Authenticating to Google
Google's Spreadsheet API uses OAuth. To easily create a Google OAuth App
1. Head to [Google's Quick Start](https://developers.google.com/sheets/api/quickstart/python#step_1_turn_on_the) and click "Enable the Google Sheets API". For OAuth Client type, select "Web Server".
2. Save `credentials.json` to your working directory
3. When you run a SheetReader instance, Google should open a browser to ask you to authenticate and authorize.

## Use Cases
- __Run spreadsheet values against API__: Our warehouse team had a spreadsheet of shipment tracking numbers. We used SheetReader to read the tracking numbers, hit the tracking API, and update the spreadsheet with delivery date.
- __Lookup spreadsheet values in database__: Our marketing team had a list of customer emails who went through a marketing campaign around a new feature. We used SheetReader to compare those emails to our production database to show which users used the new feature. todo: Add postgres example
