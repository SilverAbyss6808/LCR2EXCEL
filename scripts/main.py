
# this is gonna be the script that actually runs the program

import data_processing as dp
import excel_rw as erw
import csv
import openpyxl
import string

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
    input_csv = erw.read_input_csv(input_csv_path)
    if input_csv:
        break

# take and verify input excel path
print('\nNext, enter the path to your current cost-tracking spreadsheet.')
while True:
    input_excel_path = input(f'Path to Excel spreadsheet: ')
    # input_excel =


