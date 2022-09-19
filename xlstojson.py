#!/usr/bin/python3.10
import json, os, sys
from xlsx_reader import get_workbook, get_workbook_data

def main():
    ''' The CLI / output task. '''
    source_file = input("Enter the path to the filename -> ")
    if os.path.isfile(source_file):
        pathname = os.path.splitext(source_file)
        file_name = pathname[0].split('/')[-1]
        try:
            output_file_name = file_name + '.json'
            workbook = get_workbook(source_file)
            workbook_data = get_workbook_data(workbook)
            with open(output_file_name, 'w+', encoding="utf-8") as output_file:
                output_file.write(json.dumps(workbook_data, sort_keys=True, indent=2, separators=(",", ": ")))
                print (f"{output_file.name} was created")
        except Exception as error:
            print("some error occured")
            print(error)
            sys.exit(2)
    else:
        print ("Sorry, that was not a valid filename")

main()
