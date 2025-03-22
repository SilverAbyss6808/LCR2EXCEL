
# this file is gonna be just for file interactions

import openpyxl as opxl
import pypdf
import string
import os


pdf_data: string


class Job:
    def __init__(self, jnum: string, desc: string, pm: string, est: int, act: int):
        self.jnum = jnum
        self.desc = desc
        self.pm = pm
        self.est = est
        self.act = act


class JobRow:
    def __init__(self, column1, jobno, description, column2, pm, column5):
        self.column1 = column1
        self.jobno = jobno
        self.description = description
        self.column2 = column2
        self.pm = pm
        self.column5 = column5

    def compare_jobs(self):
        pass


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
    jobs: JobRow[]:
        pass

    reader = pypdf.PdfReader(path)
    pages = reader.pages
    for page in pages:
        data += page.extract_text()
    interpret_pdf_data(data)
    return data


def interpret_pdf_data(data: string):
