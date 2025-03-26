
# this is gonna be the script that actually runs the program

import data_processing as dp
import file_rw as frw
import openpyxl as opxl
import string
import os

input_pdf_path: string  # path to pdf file
input_excel_path: string    # path to original spreadsheet
output_excel_path: string   # path to output spreadsheet: should be in the same directory as the input

print(
    f'###############################################\n'
    f'########## WELCOME TO LCR2EXCEL v0.1 ##########\n'
    f'###############################################'
)

# take pdf path as input and make sure it exists
print('\nTo begin, enter the path to your Labor Cost Report .pdf file.')
while True:
    input_pdf_path = input(f'Path to pdf file: ')
    input_pdf = frw.read_input_file(input_pdf_path, '.pdf')
    if input_pdf:
        break

# take and verify input excel path
print('\nNext, enter the path to your current cost-tracking spreadsheet.')
while True:
    # todo: if no input spreadsheet (e.g. new year), create one
    input_excel_path = input(f'Path to Excel spreadsheet (if none, press the Enter key): ')
    input_excel = frw.read_input_file(input_excel_path, '.xlsx')
    if input_excel:
        break

# print proposed output file and path
# output_excel_path = f'{input_excel_path}-MODIFIED'
split_ext = os.path.splitext(input_excel_path)
output_excel_path = f'{split_ext[0]}-MODIFIED{split_ext[1]}'

while True:
    confirm = input(f'\nProposed output file {output_excel_path}. Is this okay? [y/n] ')
    confirm = confirm.lower()
    if confirm == 'y':
        print(f'Proceeding to file processing...')
        break
    elif confirm == 'n':
        new_path: string = input(f'Enter proposed path: ')
        new_type = os.path.splitext(new_path)[1]
        validate_path = frw.validate_proposed_filepath(new_path, new_type)
        if validate_path:
            output_excel_path = new_path
            print(f'Path {output_excel_path} is valid. Proceeding to file processing...')
            break

# process data !!
jobs_from_pdf = frw.read_pdf(input_pdf_path)  # returns the relevant info from the pdf
jobs_from_excel = frw.read_excel(input_excel_path)
