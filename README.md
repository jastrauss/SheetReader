# SheetReader

Iterate through rows of Google Spreadsheet as if it were a Python CSV DictReader or DictWriter.

## Getting Started
For a full example, see [example.py](example.py) or watch a quick Youtube [demo](https://youtu.be/0fdjtsJqVXI)

        reader = SheetReader(
            workbook_id = 'workbook_id_123', 
            sheet_name = 'Sheet1', 
            header_range=('A1', 'D1'
            data_range = ('A1','D10')
        )

        for row in reader:
            name = row["name"]
            print "Hello", name
            row["Reversed Name"] = name[::-1]
