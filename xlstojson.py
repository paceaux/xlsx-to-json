#!/usr/bin/python3.10
import json
import os.path
import pylightxl as xl

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
        column_name = column_names[counter]
        #TODO: this doesn't format any for a cell. Consider formatting date/numbers
        row_data[column_name] = cell
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

def main():
    ''' The CLI / output task. '''
    filename = input("Enter the path to the filename -> ")
    if os.path.isfile(filename):
        workbook = get_workbook(filename)
        workbook_data = get_workbook_data(workbook)
        output = \
		open((filename.replace("xlsx", "json")).replace("xls", "json"), "w+", encoding="utf-8")
        output.write(json.dumps(workbook_data, sort_keys=True, indent=2,  separators=(',', ": ")))
        output.close()
        print (f"{output.name} was created")
    else:
        print ("Sorry, that was not a valid filename")

main()
