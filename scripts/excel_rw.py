
# so what i wanna do is read each cell and make an array, then edit the ones that need editing,
# and then write them back. basically

# this file is gonna be just for excel sheet interactions


import openpyxl
import string
import csv
import os


csv_data: csv


def read_input_file(path: string, filetype: string):
    # check if the path exists and is a valid csv/xlsx
    try:

        if os.path.splitext(path)[1] != filetype:
            raise TypeError(f'not a {filetype} file')

        if filetype == '.csv':
            print(f'{type}')
            # file_in = open(path, 'r')
            # reader = csv.reader(file_in)
            # TODO: add to external variable to allow access later??
        elif filetype == '.xlsx':
            print(f'{type}')
        else:
            raise TypeError(f'not a file of type .csv or .xlsx')

        return True  # returns True if the file exists

    except ValueError as ve:  # file's not open probably
        print(f'Error: {ve}')
    except NameError as ne:  # undefined variables, etc
        print(f'Error: {ne}')
    except FileNotFoundError as nf:  # no file at specified path
        print(f'Error: {nf}')
    except TypeError as te:  # provided file is not csv or xlsx, depending on type
        print(f'Error: {te}')

    return False  # only executes if an error occurs


def validate_filepath(path: string, filetype: string):
    directory = os.path.dirname(path)
    if directory is not None:
        if not os.path.isfile(path):
            print(f'File will be created at {path}.')
            return True
        else:
            print('Directory is valid, but file already exists. Please enter a path with a new filename.')
            return False
    else:
        print(f'Path {path} is not in a valid directory.')
        return False

