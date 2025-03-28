
# this file is gonna be just for file interactions

import openpyxl as opxl
import pypdf
import string
import os
import data_processing as dp
import visual_formatting as vf


pdf_date: string


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

    rows: list = list(active_sheet.iter_rows(min_row=2, max_col=active_sheet.max_column, values_only=True))
    num_rows = active_sheet.max_row - 13  # starting from line 2, run until 12 lines from end cuz those aren't jobs
    num_jobs: int = int(num_rows / 4)  # number of lines minus title line, divided by four lines per job

    for i in range(0, num_rows):  # i is the index jsyk
        current_row = (rows[i])
        excel_jobs.append(current_row)

    excel_jobs = dp.create_jobs_from_excel_in(excel_jobs, active_sheet.max_column)
    return excel_jobs


def get_title_row(input_excel_path: string):
    workbook = opxl.load_workbook(input_excel_path)
    active_sheet = workbook.active

    row: list[str] = []

    for i in range(1, active_sheet.max_column):
        value = active_sheet.cell(1, i).value
        row.append(value)

    return row


def create_write_new_excel(new: list[dp.Job], old: list[dp.Job], old_path: string, new_path: string):
    if old_path != '':
        job_list = dp.compare_jobs(new, old)
        max_col = opxl.load_workbook(old_path).active.max_column
        title_row = get_title_row(old_path)
    else:
        job_list = new
        max_col = 6
        title_row = 'Column1', 'Job No', 'Description', 'Column2', 'PM', 'Column5'

    formatted_job_list = dp.format_jobs_as_excel(job_list, max_col)

    new_file = opxl.Workbook()
    sheet = new_file.active

    sheet.append(title_row)

    for row in formatted_job_list:
        sheet.append(row)

    # todo: append end stuff

    # formatting :3
    vf.format_widths(sheet)
    vf.color_every_other_line(sheet, '00FFFFFF', '00DDEBF7')


    new_file.save(new_path)

    return formatted_job_list
