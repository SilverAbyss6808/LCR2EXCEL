
# so what i wanna do is read each cell and make an array, then edit the ones that need editing,
# and then write them back. basically

# this file is gonna be just for excel sheet interactions


import openpyxl
import string
import csv


csv_data: csv

def read_input_csv(path: string):
    # check if the path exists and is a valid csv
    try:
        csv_in = open(path, 'r')
        reader = csv.reader(csv_in)
        return True # returns True if the csv file exists
    except ValueError as ve:
        # file's not open probably
        print(f'Error: {ve}')
    except NameError as ne:
        # undefined variables, etc
        print(f'Error: {ne}')
    except FileNotFoundError as nf:
        print(f'Error: {nf}')

    return False    # only executes if an error occurs
