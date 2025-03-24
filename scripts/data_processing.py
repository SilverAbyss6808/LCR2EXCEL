# this is where the stuff for processing the input pdf is gonna go
# plus any other data processing things

import openpyxl as opxl
import string


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


def process_data(input_pdf_path: string, input_excel_path: string, output_file_path: string):
    workbook = opxl.load_workbook(input_excel_path)
    active_sheet = workbook.active


def format_pdf_data_as_job(data: string):  # return an array of jobs
    num_jobs: int = 0
    data_lines = data.split('\n')

    for index, line in enumerate(data_lines):
        if 'Job Totals' in line:
            if 'Primary' not in line:
                line += '\n'
                num_jobs += 1

    jobs = create_jobs_from_raw(data_lines, num_jobs)

    return jobs


def create_jobs_from_raw(data: list[string], num_jobs: int):
    # create jobs from split string data
    jobs: list[Job] = []

    # setting some recognizable defaults so i know if somethings wrong
    jnum: string = 'DEFAULT'
    desc: string = 'DEFAULT'
    pm: string = 'DEFAULT'
    # est: int = 99999999
    # act: int = 99999999
    est: string = 'DEFAULT'
    act: string = 'DEFAULT'

    cc = False

    for line in data:
        if line == 'Cost Code Description':
            cc = True
        elif cc:
            jnum = line.split(' ')[0]
            desc = line.replace(jnum + ' ', '')
            cc = False
        elif 'Est Actual Remaining' in line:
            pm = line.split(' ')[0]
        elif 'Job Totals' in line and 'Primary' not in line:
            # todo: split em right heehooooooo

            line_nospaces = line.replace(' ', '')

            # est = line.split(' ')[4]
            # act = line.split(' ')[5]

            est = line_nospaces.split('*')[3]
            act = line_nospaces.split('*')[4]
        else:
            continue

        if jnum != 'DEFAULT' and desc != 'DEFAULT' and pm != 'DEFAULT' and est != 'DEFAULT' and act != 'DEFAULT':
            # add job to list and reset variables
            jobs.append(Job(jnum, desc, pm, est, act))

            jnum = 'DEFAULT'
            desc = 'DEFAULT'
            pm = 'DEFAULT'
            est = 'DEFAULT'
            act = 'DEFAULT'

    # reports if there's a discrepancy found so the data can be reviewed
    if len(jobs) != num_jobs:
        print(f'Though {num_jobs} were found, {len(jobs)} were actually reported. You may want to check your data.')

    return jobs

# process_data('..\\io\\testfile.pdf',
#              '..\\io\\Labor Tracking Spreadsheet 2024.xlsx',
#              '..\\io\\Labor Tracking Spreadsheet 2024-MODIFIED.xlsx')
