
# this is gonna be the script that actually runs the program

import data_processing as dp
import excel_rw as erw
import csv
import openpyxl
import string
import os

input_csv_path: string  # path to csv file
input_excel_path: string    # path to original spreadsheet
output_excel_path: string   # path to output spreadsheet: should be in the same directory as the input

print(
    f'###############################################\n'
    f'########## WELCOME TO LCR2EXCEL v0.1 ##########\n'
    f'###############################################'
)

# take csv path as input and make sure it exists
print('\nTo begin, enter the path to your Labor Cost Report .csv file.')
while True:
    input_csv_path = input(f'Path to csv file: ')
    input_csv = erw.read_input_file(input_csv_path, '.csv')
    if input_csv:
        break

# take and verify input excel path
print('\nNext, enter the path to your current cost-tracking spreadsheet.')
while True:
    input_excel_path = input(f'Path to Excel spreadsheet: ')
    input_excel = erw.read_input_file(input_excel_path, '.xlsx')
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
        validate_path = erw.validate_filepath(new_path, '.xlsx')
        if validate_path:
            output_excel_path = new_path
            print(f'Path {output_excel_path} is valid. Proceeding to file processing...')
            break


