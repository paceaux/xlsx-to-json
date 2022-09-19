'''Module that reads an xlsx spreadsheet and can produce json data from it'''
import pylightxl as xl
import datetime
from dateutil.parser import parse

def is_date(string, fuzzy=False):
    """
    Return whether the string can be interpreted as a date.

    :param string: str, string to check for date
    :param fuzzy: bool, ignore unknown tokens in string if True
    https://stackoverflow.com/a/25341965/1045901
    """
    try:
        parse(string, fuzzy=fuzzy)
        return True

    except ValueError:
        return False

def get_column_names(sheet):
    '''Takes a single worksheet, returns the strings in the top row of each column'''
    column_lists = sheet.cols
    column_names = []

    for column_list in column_lists:
        column_names.append(column_list[00])

    return column_names

def get_row_data(row, column_names):
    '''takes a single row of a worksheet and an array of rows,
        returns an object with column_name:rowvalue
    '''
    row_data = {}
    counter = 0

    for cell in row:
        column_name = column_names[counter].lower().replace(' ', '_')
        cell_value = cell
        if is_date(cell):
            cell_value = datetime.datetime(cell_value).isoformat()
        row_data[column_name] = cell_value
        counter = counter + 1

    return row_data

def get_sheet_data(sheet, column_names):
    '''Takes a single worksheet, returns an object with row data'''
    max_rows = sheet.size[0]
    sheet_data = []

    for idx in range(2, max_rows):
        row = sheet.row(idx)
        row_data = get_row_data(row, column_names)
        sheet_data.append(row_data)

    return sheet_data

def get_workbook_data(workbook):
    '''Takes a workbook and returns all worksheet data'''
    workbook_sheet_names = workbook.ws_names
    workbook_data = {}

    for sheet_name in workbook_sheet_names:
        worksheet = workbook.ws(ws=sheet_name)
        column_names = get_column_names(worksheet)
        sheet_data = get_sheet_data(worksheet, column_names)
        workbook_data[sheet_name.lower().replace(' ', '_')] = sheet_data

    return workbook_data

def get_workbook(filename):
    '''opens a workbook for reading'''
    return xl.readxl(filename)
