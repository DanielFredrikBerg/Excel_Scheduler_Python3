#!/usr/bin/python3

from openpyxl import Workbook
import openpyxl
import schedlr
import json
import os










def initialize_excel_file(filename='7_19.xlsx'):
    
    wb.save(filename=filename)


def generate_xlsx_color_scheme(file_, start_cell):
    for i in range(2, 5): 
        ws.merge_cells(start_row=i, start_column='F', end_row=5, end_column='I')
    
    



def main(filename = "tdp007_019_assistent_schema.xlsx", csv="tdp007_019.csv"):
    # Get correctly formated schedule info.
    formatted_csv = "7_19.csv"
    xl = '7_19.xlsx' 
    schedlr.format_timeedit_categories(csv, formatted_csv)
    schedule_dict = schedlr.initialize_week_schedule(formatted_csv)
    #print(json.dumps(schedule_dict, indent=2))

    # Create Excel file
    assistants = schedlr.get_assistants(formatted_csv, 'TDP007') 
    #print(assistants)


    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Schema tdp007 och tdp019"

    for i in range(2, 5): 
        ws1.merge_cells('F2:I2')
#        ws1.merge_cells(start_row=i, start_column='F', end_row=5, end_column='I')

    wb.save(filename=xl)
    #initialize_excel_file()
    
    return "0"



if __name__ == '__main__':
    main()
else:
    print("exceller running as module")


