#!/usr/bin/python3.10
import json
import os.path
from xlsx_reader import get_workbook, get_workbook_data

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
