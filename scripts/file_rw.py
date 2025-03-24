
# this file is gonna be just for file interactions

import openpyxl as opxl
import pypdf
import string
import os
import data_processing as dp


pdf_data: string


def read_input_file(path: string, filetype: string):
    # check if the path exists and is a valid pdf/xlsx
    try:
        if os.path.splitext(path)[1] != filetype:
            raise TypeError(f'not a {filetype} file')

        # todo: verify that input files exist !!!
        if filetype == '.pdf':
            read_pdf(path)
        elif filetype == '.xlsx':
            pass
        else:
            raise TypeError(f'not a file of type .pdf or .xlsx')

        return True  # returns True if the file exists

    except ValueError as ve:  # file's not open probably
        print(f'Error: {ve}')
    except NameError as ne:  # undefined variables, etc
        print(f'Error: {ne}')
    except FileNotFoundError as nf:  # no file at specified path
        print(f'Error: {nf}')
    except TypeError as te:  # provided file is not pdf or xlsx, depending on type
        print(f'Error: {te}')

    return False  # only executes if an error occurs


def validate_filepath(path: string, filetype: string):
    directory = os.path.dirname(path)
    if directory is not None:
        if not os.path.isfile(path):
            if filetype == '.xlsx':
                return True
        else:
            # todo: ask if it's ok to overwrite the file rather than just saying fuck you
            print('Directory is valid, but file already exists. Please enter a path with a new filename.')
            return False
    else:
        print(f'Path {path} is not in a valid directory.')
        return False


def read_pdf(path: string):
    # path has already been verified so its ok
    data: string = ''

    reader = pypdf.PdfReader(path)
    pages = reader.pages
    for page in pages:
        data += page.extract_text()
        print(page.extract_text())

    jobs_list: list[dp.Job] = dp.format_pdf_data_as_job(data)
    return jobs_list


jobs = read_pdf('..\\io\\Tuttle Labor Cost.pdf')
for job in jobs:
    print(f'JobNum: {job.jnum}, Desc: {job.desc}, PM: {job.pm}, Est: {job.est}, Act: {job.act}')
