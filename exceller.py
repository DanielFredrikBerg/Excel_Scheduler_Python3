#!/usr/bin/python3

from openpyxl import Workbook
import schedlr
import json











def initialize_excel_file():
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Schema tdp007 och tdp019"
    wb.save(filename='tdp007.xlsx')




def main(filename = "tdp007_019_assistent_schema.xlsx", csv="tdp007_019.csv"):
    # Get correctly formated schedule info.
    schedlr.format_timeedit_categories(csv, "o_csv.csv")
    schedule_dict = schedlr.initialize_week_schedule("o_csv.csv")
    print(json.dumps(schedule_dict, indent=2))

    # Create Excel file

    initialize_excel_file()
    
    return "0"



if __name__ == '__main__':
    main()
else:
    print("exceller running as module")


