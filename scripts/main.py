
# this is gonna be the script that actually runs the program

import file_rw as frw
import string
import os
import main_choices as choice


input_pdf_path: string  # path to pdf file
input_excel_path: string    # path to original spreadsheet
output_excel_path: string  # path to output spreadsheet: should be in the same directory as the input


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
    input_excel_path = input(f'Path to Excel spreadsheet (if none, press the Enter key): ')
    if input_excel_path != '':
        input_excel = frw.read_input_file(input_excel_path, '.xlsx')
        if input_excel:
            break
    else:
        input_excel_path = ''
        break

# print proposed output file and path
if input_excel_path != '':
    split_ext = os.path.splitext(input_excel_path)
    output_excel_path = f'{split_ext[0]}-MODIFIED{split_ext[1]}'
else:
    output_excel_path = f'NewSpreadsheet.xlsx'

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
try:
    if input_excel_path == '':
        choice.create_new_excel_from_pdf(input_pdf_path, output_excel_path)
    else:
        choice.add_pdf_data_to_existing_spreadsheet(input_pdf_path, input_excel_path, output_excel_path)
except:
    print('An error occurred. Please try again.')
    exit(1)

exit_confirm = input(f'{output_excel_path} successfully created! Press Enter to exit.')
exit(0)
