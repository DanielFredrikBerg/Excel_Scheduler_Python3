#!/usr/bin/python3

from openpyxl import Workbook
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
import schedlr
import json
import os










def initialize_excel_file(filename='7_19.xlsx'):    
    wb.save(filename=filename)


def merge_row(ws_, row, start_col, end_col):
    ws_.merge_cells(f'{start_col}{row}:{end_col}{row}')

    
def merge_rows(ws_, start_row, end_row, start_col, end_col):
    for row in range(start_row, end_row + 1):
        merge_row(ws_, start_row, start_col, end_col)


def merge_area(ws_, start_row, end_row, start_col, end_col):
    for row in range(start_row, end_row + 1):
        ws_.merge_cells(f'{start_col}{start_row}:{end_col}{end_row}')
    
# top_left_cell.fill = PatternFill("solid", fgColor="DDDDDD")
def color_fill(cell, color):
    cell.fill = PatternFill("solid", fgColor=color)


def paint_area(ws_, color, start_row, end_row, start_col, end_col):
    columns = list(map(chr, range(ord(start_col), ord(end_col)+1)))
    for row in range(start_row, end_row + 1):
        for col in columns:
            color_fill(ws_[f"{col}{row}"], color)

            
def paint_row(ws_, color, start_row, start_col, end_col):
    columns = list(map(chr, range(ord(start_col), ord(end_col)+1)))
    for col in columns:
        color_fill(ws_[f"{col}{start_row}"], color)
        

def generate_color_instructions(ws_, start_row, start_col, end_col):
    color = get_color_scheme()
    colors = ['green', 'yellow', 'red']
    info_text = ['Tillgänglig', 'Krock med icke obligatoriskt moment', 'Krock med obligatoriskt moment']
    info_counter = 0
    row_counter = start_row
    for color_ in colors:
        paint_row(ws_, color[color_], row_counter, start_col, end_col)
        ws_.cell(row=row_counter, column=start_col).value = info_text[info_counter] # set value to cell coordinates
        
        merge_row(ws_, row_counter, start_col, chr(ord(end_col)-1))
        row_counter += 1
        info_counter += 1
            
    
def get_color_scheme():
    colors = {
        'green' : 'b6d7a8',
        'yellow' : 'f9cb9c',
        'red' : 'ea9999',
        'lightgreen' : 'e2efda',
        'lightyellow' : 'fff2cc',
        'lightred' : 'fce4d6',
        'lightblue' : 'd9e1f2'        
    }
    return colors


def main(filename = "tdp007_019_assistent_schema.xlsx", csv="tdp007_019.csv"):
    # Get correctly formated schedule info.
    formatted_csv = "7_19.csv"
    xl = '7_19.xlsx' 
    schedlr.format_timeedit_categories(csv, formatted_csv)
    schedule_dict = schedlr.initialize_week_schedule(formatted_csv)
    #print(json.dumps(schedule_dict, indent=2))

    # Create Excel file
    assistants = schedlr.get_assistants(formatted_csv, 'TDP007') 

    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Schema tdp007 och tdp019"


    generate_color_instructions(ws1, 2, 'F', 'J')

    
    wb.save(filename=xl)
    #initialize_excel_file()
    
    return "0"



if __name__ == '__main__':
    main()
else:
    print("exceller running as module")


