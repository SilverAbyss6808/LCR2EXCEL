
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

        if os.path.exists(path):
            if filetype == '.pdf':
                read_pdf(path)
            elif filetype == '.xlsx':
                pass
        else:
            raise FileNotFoundError(f'File {path} does not exist. Please try again.')

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


def validate_proposed_filepath(path: string, filetype: string):
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

    jobs_list: list[dp.Job] = dp.format_pdf_data_as_job(data)
    return jobs_list


def read_excel(input_excel_path: string):
    excel_jobs: list[string] = []

    workbook = opxl.load_workbook(input_excel_path)
    active_sheet = workbook.active

    rows: list = list(active_sheet.iter_rows(min_row=2, max_col=4, values_only=True))
    num_rows = active_sheet.max_row - 13  # starting from line 2, run until 12 lines from end cuz those aren't jobs
    num_jobs: int = int(num_rows / 4)  # number of lines minus title line, divided by four lines per job

    for i in range(0, num_rows, 4):  # i is the index jsyk
        current_row = (rows[i])
        excel_jobs.append(current_row)

    excel_jobs = dp.create_jobs_from_excel_in(excel_jobs)
    return excel_jobs


def write_new_excel(new: list[dp.Job], old: list[dp.Job]):
    job_list = dp.compare_jobs(new, old)
    for i in job_list:
        print(str(i))


# uncomment for test
# both return lists of jobs !!! with relevant info !!!!!
# except for the stuff from the original excel, those dont have costs, those will be edited when the two are merged
new_jobs = read_pdf('..\\io\\Tuttle Labor Cost.pdf')
orig_jobs = read_excel('..\\io\\Labor Tracking Spreadsheet 2024.xlsx')

for job in new_jobs:
    print(f'NEW JOB: {str(job)}')

for job in orig_jobs:
    print(f'OLD JOB: {str(job)}')

print(f'Number of new jobs: {len(new_jobs)}\n'
      f'Number of preexisting jobs: {len(orig_jobs)}'
      f'Number of combined jobs: ')

write_new_excel(new_jobs, orig_jobs)
