import re
from string import ascii_letters

from google_sheets_api import GoogleSheetsConnector

""" A1 => 1 """
def get_row_from_cell(cell):
    return int(re.split('[a-zA-Z]+', cell)[1])

""" A1 => A """
def get_col_from_cell(cell):
    return re.split('\d+', cell)[0]

""" 0 => A """
def number_to_excel_column(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

""" A => 0 """
def excel_column_to_number(col):
    num = 0
    for c in col:
        if c in ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

class SheetReader:
    """
    Represents a connection to a Google Sheet. Returns an iterator of Row objects
    which provide access to each row at a time.
    """

    def __init__(self, workbook_id, sheet_name, header_range=[], data_range=[], read_chunk_size=100, write_chunk_size=100, auto_update=True):

        self.workbook_id = workbook_id
        self.sheet_name = sheet_name

        self.header_start_cell = header_range[0]
        self.header_end_cell = header_range[1]
        self.data_start_cell = data_range[0]
        self.data_end_cell = data_range[1]
        self.data_start_col = get_col_from_cell(self.data_start_cell)
        self.data_end_col = get_col_from_cell(self.data_end_cell)

        self.read_chunk_size = read_chunk_size
        self.write_chunk_size = write_chunk_size
        self.write_map = {}
        self.auto_update = auto_update

        start_row = get_row_from_cell(self.data_start_cell)

        # Internal State for read chunks
        self.current_read_chunk = []
        self.read_chunk_start = None # row
        self.read_chunk_end = start_row - 1 # row
        self.current_row_index = start_row - 1 # current iteration

        self.connection = GoogleSheetsConnector(workbook_id, sheet_name)

        self.headers = []
        self.header_map = {} # {'name':0, 'age':1, ... }, offset from first column in the dataset
        self.get_headers()


        if len(data_range) < 2:
            raise ValueError('SheetReader data_range must contain at a start and end cell')

        if len(header_range) < 2:
            raise ValueError('SheetReader header_range must contain at a start and end cell')

        header_start_row = get_row_from_cell(self.header_start_cell)
        header_end_row = get_row_from_cell(self.header_end_cell)
        if header_start_row != header_end_row:
            raise ValueError('SheetReader header_range must be one row')

    def get_headers(self):
        raw_values = self.connection.read_range(self.header_start_cell, self.header_end_cell)
        self.headers = raw_values[0]

        # Create a map from header name to it's column index, eg. 'Name' => Column 0
        # We use this map to access elements in a row by their header name, eg. Row.get('Name') => Jim
        self.header_map = {header: index for (index, header) in enumerate(self.headers)}

    def update(self):
        '''
            Rather than make a request for each write, we save them in memory
            and write them in chunks of write_chunk_size
        '''
        self.connection.bulk_write_range(self.write_map)
        self.write_map = {}


    def get_row_values(self, row_index):
        '''
            Rather than make a request to the spreadsheet for each row,
            save a chunk in memory of size read_chunk_size
            and read more when we run out
        '''
        if row_index > self.read_chunk_end: # we don't have the row in memory
            new_chunk_start = self.read_chunk_end + 1
            new_chunk_end = self.read_chunk_end + self.read_chunk_size
            if (new_chunk_end) > get_row_from_cell(self.data_end_cell): # don't shoot past the desired range
                new_chunk_end = self.read_chunk_end + self.read_chunk_end

            # get the chunk
            new_chunk_start_cell = self.data_start_col + str(new_chunk_start)
            new_chunk_end_cell = self.data_end_col + str(new_chunk_end)
            self.current_read_chunk = self.connection.read_range(new_chunk_start_cell, new_chunk_end_cell)

            # update our indexes
            self.read_chunk_start = new_chunk_start
            self.read_chunk_end = new_chunk_end

            return self.current_read_chunk[row_index - self.read_chunk_start]

        elif (row_index >= self.read_chunk_start) and (row_index <= self.read_chunk_end):
            # todo: check if the range the specified is too big
            # except IndexError:
            return self.current_read_chunk[row_index - self.read_chunk_start]


    def __iter__(self):
        return self

    def __next__(self):
        if self.current_row_index > get_row_from_cell(self.data_end_cell) - 1:
            raise StopIteration
        else:
            self.current_row_index += 1
            row_values = self.get_row_values(self.current_row_index)
            data_start_cell = self.data_start_col + str(self.current_row_index)
            data_end_cell = self.data_end_col + str(self.current_row_index)
            data_range = (data_start_cell, data_end_cell)
            return Row(self, data_range, row_values)

        next = __next__ # Python 2 iterators look for "next"

    def __del__(self):
        if bool(self.write_map):
            self.update()

class Row:
    """
    A dict like object that represent one row in a Google sheet.

    Read a value: row[column_name]
    Write a value: row[column_name] = new_value
    """

    def __init__(self, sheet_reader_instance, data_range, values):

        self.sheet_reader_instance = sheet_reader_instance
        self.workbook_id = sheet_reader_instance.workbook_id
        self.sheet_name = sheet_reader_instance.sheet_name
        self.write_map = sheet_reader_instance.write_map

        self.data_start_cell = data_range[0]
        # self.data_end_cell = data_range[1]
        self.current_row_index = get_row_from_cell(self.data_start_cell)
        self.id = get_row_from_cell(self.data_start_cell)
        self.header_map = sheet_reader_instance.header_map
        self.values = values

    def __getitem__(self, key):
        # todo: throw error if doesn't exist
        field_index = self.header_map.get(key, '___') # todo, this used to be None, but would fail the next check on col '0'
        # if not field_index:
        if field_index == '___':
            raise KeyError
        return self.values[field_index]

    def __setitem__(self, key, value, immediate_update=False):
        # todo: accept a dictionary or **kwargs
        cell_col = self.header_map.get(key, None)
        if not cell_col:
            raise KeyError
        # offset by the range of the dataset start cell
        data_start_col = get_col_from_cell(self.data_start_cell) # returns a letter
        data_start_col = excel_column_to_number(data_start_col)
        cell_col = data_start_col + cell_col
        cell_col = number_to_excel_column(cell_col)
        # convert our cell column back to a letter

        destination_cell = cell_col + str(self.current_row_index)
        # value = [[value]]

        # print 'Set cell %s to "%s"' % (destination_cell, value)
        # We try to group writes (to minimize https requests)
        if immediate_update:
            self.sheet_reader_instance.connection.write_range(destination_cell, destination_cell, [[value]])
            return

        self.write_map[destination_cell] = value
        if not self.sheet_reader_instance.auto_update:
            return

        # let's make a request every 100 cell writes
        if len(self.write_map.keys()) >= self.sheet_reader_instance.write_chunk_size:
            self.sheet_reader_instance.update()

    def __str__(self):

        # return ','.join(self.values)

        pretty_dict = {}
        for k, v in self.header_map.items():
            pretty_dict[k] = self.values[v]

        return str(pretty_dict)
