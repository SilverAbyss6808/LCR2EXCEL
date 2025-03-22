
# this is where the stuff for processing the input pdf is gonna go
# plus any other data processing things

import openpyxl as opxl
import file_rw as frw
import string


pdf_data: string


def process_data(input_pdf_path: string, input_excel_path: string, output_file_path: string):
    workbook = opxl.load_workbook(input_excel_path)
    active_sheet = workbook.active


def format_pdf_data():
    pass


process_data('..\\io\\testfile.pdf',
             '..\\io\\Labor Tracking Spreadsheet 2024.xlsx',
             '..\\io\\Labor Tracking Spreadsheet 2024-MODIFIED.xlsx')
