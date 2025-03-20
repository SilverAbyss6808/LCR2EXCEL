
# this is where the stuff for processing the input csv is gonna go
# plus any other data processing things

import openpyxl as opxl
import excel_rw as erw
import string

def process_data(input_csv_path: string, input_excel_path: string, output_file_path: string):
    workbook = opxl.load_workbook(input_excel_path)
    active_sheet = workbook.active


process_data('..\\io\\testfile.csv',
             '..\\io\\Labor Tracking Spreadsheet 2024.xlsx',
             '..\\io\\Labor Tracking Spreadsheet 2024-MODIFIED.xlsx')
