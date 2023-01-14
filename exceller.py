#!/usr/bin/python3

from openpyxl import Workbook
import schedlr
import json
import os










def initialize_excel_file(filename='7_19.xlsx'):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Schema tdp007 och tdp019"
    wb.save(filename=filename)




def main(filename = "tdp007_019_assistent_schema.xlsx", csv="tdp007_019.csv"):
    # Get correctly formated schedule info.
    formatted_csv = "7_19.csv"
    schedlr.format_timeedit_categories(csv, formatted_csv)
    schedule_dict = schedlr.initialize_week_schedule(formatted_csv)
    #print(json.dumps(schedule_dict, indent=2))

    # Create Excel file
    assistants = schedlr.get_assistants(formatted_csv, 'TDP007') 
    #print(assistants)
    initialize_excel_file()
    
    return "0"



if __name__ == '__main__':
    main()
else:
    print("exceller running as module")


