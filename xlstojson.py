#!/usr/bin/python2.7
import pylightxl as xl
import json
import os.path
import datetime

def get_column_names(sheet):
	column_lists = sheet.cols
	column_names = []
 
	for columnList in column_lists:
		column_names.append(columnList[00])

	return column_names

def get_row_data(row, column_names):
	row_data = {}
	counter = 0

	for cell in row:
		column_name = column_names[counter]
		row_data[column_name] = cell
		counter = counter + 1

	return row_data

def get_sheet_data(sheet, column_names):
	max_rows = sheet.size[0]
	sheet_data = []

	for idx in range(2, max_rows):
		row = sheet.row(idx)
		rowData = get_row_data(row, column_names)
		sheet_data.append(rowData)

	return sheet_data

def get_workbook_data(workbook):
	workbook_sheet_names = workbook.ws_names
	workbook_data = {}

	for sheet_name in workbook_sheet_names:
		worksheet = workbook.ws(ws=sheet_name)
		column_names = get_column_names(worksheet)
		sheet_data = get_sheet_data(worksheet, column_names)
		workbook_data[sheet_name.lower().replace(' ', '_')] = sheet_data

	return workbook_data

def get_workbook(filename):
	return xl.readxl(filename)

def main():
	filename = input("Enter the path to the filename -> ")
	if os.path.isfile(filename):
		workbook = get_workbook(filename)
		workbook_data = get_workbook_data(workbook)
		output = \
		open((filename.replace("xlsx", "json")).replace("xls", "json"), "w+")
		output.write(json.dumps(workbook_data, sort_keys=True, indent=2,  separators=(',', ": ")))
		output.close()
		print ("%s was created" %output.name)
	else:
		print ("Sorry, that was not a valid filename")

main()
