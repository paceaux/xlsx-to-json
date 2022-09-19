#!/usr/bin/python3.10
'''Module that reads an xls spreadsheet and can produce json data from it'''
import datetime
import xlrd

def get_column_names(sheet):
    '''Takes a single worksheet, returns the strings in the top row of each column'''
    row_size = sheet.row_len(0)
    col_values = sheet.row_values(0, 0, row_size )
    column_names = []

    for value in col_values:
        column_names.append(value)

    return column_names

def get_row_data(row, column_names):
    '''takes a single row of a worksheet and an array of rows,
        returns an object with column_name:rowvalue
    '''
    row_data = {}
    counter = 0

    for cell in row:
        # check if it is of date type print in iso format
        if cell.ctype==xlrd.XL_CELL_DATE:
            row_data[column_names[counter].lower().replace(' ', '_')] = datetime.datetime(*xlrd.xldate_as_tuple(cell.value,0)).isoformat()
        else:
            row_data[column_names[counter].lower().replace(' ', '_')] = cell.value
        counter +=1

    return row_data

def get_sheet_data(sheet, column_names):
    '''Takes a single worksheet, returns an object with row data'''
    num_of_rows = sheet.nrows
    sheet_data = []

    for idx in range(1, num_of_rows):
        row = sheet.row(idx)
        row_data = get_row_data(row, column_names)
        sheet_data.append(row_data)

    return sheet_data

def get_workbook_data(workbook):
    '''Takes a workbook and returns all worksheet data'''
    nsheets = workbook.nsheets
    workbook_data = {}

    for idx in range(0, nsheets):
        worksheet = workbook.sheet_by_index(idx)
        column_names = get_column_names(worksheet)
        sheetdata = get_sheet_data(worksheet, column_names)
        workbook_data[worksheet.name.lower().replace(' ', '_')] = sheetdata

    return workbook_data

def get_workbook(filename):
    '''opens a workbook for reading'''
    return xlrd.open_workbook(filename)
