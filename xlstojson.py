#!/usr/bin/python2.7
import pylightxl as xl
import json
import os.path
import datetime

def getColNames(sheet):
	columnLists = sheet.cols
	columnNames = []
 
	for columnList in columnLists:
		columnNames.append(columnList[00])

	return columnNames

def getRowData(row, columnNames):
	rowData = {}
	counter = 0

	for cell in row:
		columnName = columnNames[counter]
		rowData[columnName] = cell
		counter = counter + 1

	return rowData

def getSheetData(sheet, columnNames):
	maxRows = sheet.size[0]
	columnLists = sheet.cols
	sheetData = []

	for idx in range(2, maxRows):
		row = sheet.row(idx)
		rowData = getRowData(row, columnNames)
		sheetData.append(rowData)

	return sheetData

def getWorkBookData(workbook):
	workbookSheetNames = workbook.ws_names
	counter = 0
	workbookdata = {}

	for sheetName in workbookSheetNames:
		worksheet = workbook.ws(ws=sheetName)
		columnNames = getColNames(worksheet)
		sheetdata = getSheetData(worksheet, columnNames)
		workbookdata[sheetName.lower().replace(' ', '_')] = sheetdata

	return workbookdata

def getWorkbook(filename):
	return xl.readxl(filename)

def main():
	filename = input("Enter the path to the filename -> ")
	if os.path.isfile(filename):
		workbook = getWorkbook(filename)
		workbookdata = getWorkBookData(workbook)
		output = \
		open((filename.replace("xlsx", "json")).replace("xls", "json"), "w+")
		output.write(json.dumps(workbookdata, sort_keys=True, indent=2,  separators=(',', ": ")))
		output.close()
		print ("%s was created" %output.name)
	else:
		print ("Sorry, that was not a valid filename")

main()
