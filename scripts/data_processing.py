
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

    # create individual jobs from string data
    for i in range(0, num_jobs):
        # setting some recognizable defaults so i know if somethings wrong
        jnum: string = 'DEFAULT'
        desc: string = 'DEFAULT'
        pm: string = 'DEFAULT'
        # est: int = 99999999
        # act: int = 99999999
        est: string = 'DEFAULT'
        act: string = 'DEFAULT'

        cc = False

        # todo: start here next. figure out how to successfully remove line from data
        for line in data:
            if 'Cost Code' in line:
                cc = True
                break
                # print('cost code in line')
            if cc:
                jnum = line.split(' ')[0]
                desc = line
                cc = False
                print('jnum and desc set')
            if 'Est Actual Remaining' in line:
                pm = line.split(' ')[0]
                print('pm set')
            if 'Job Totals' in line:
                est = line.split(' ')[4]
                act = line.split(' ')[5]
                print('est/act set')
                break

        jobs.append(Job(jnum, desc, pm, est, act))

    print(f'{len(jobs)}/{num_jobs} jobs total.')
    return jobs


# process_data('..\\io\\testfile.pdf',
#              '..\\io\\Labor Tracking Spreadsheet 2024.xlsx',
#              '..\\io\\Labor Tracking Spreadsheet 2024-MODIFIED.xlsx')