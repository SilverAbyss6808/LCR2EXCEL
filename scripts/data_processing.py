# this is where the stuff for processing the input pdf is gonna go
# plus any other data processing things

import openpyxl as opxl
import string


class Job:
    jnum: int  # will need to cast to string with - for excel sheet (XX-XX-XXXX)
    desc: string
    pm: string
    est: int
    act: int

    def __init__(self, jnum: int, desc: string, pm: string, est: int, act: int):
        self.jnum = jnum
        self.desc = desc
        self.pm = pm
        self.est = est
        self.act = act

    def __str__(self):
        return (f'Job Number: {self.jnum}, Project Manager: {self.pm}, Description: {self.desc}\n'
                f'Estimated Cost: {self.est}, Actual Cost: {self.act}\n')


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
    jnum: int = 99999999
    desc: string = 'DEFAULT'
    pm: string = 'DEFAULT'
    est: int = 99999999
    act: int = 99999999

    cc = False

    for line in data:
        if line == 'Cost Code Description':
            cc = True
        elif cc:
            jnum_string = line.split(' ')[0]
            jnum = int(jnum_string.replace('-', ''))

            desc = line.replace(jnum_string + ' ', '')
            cc = False
        elif 'Est Actual Remaining' in line:
            pm = line.split(' ')[0]
        elif 'Job Totals' in line and 'Primary' not in line:
            line_nospaces = line.replace(' ', '')

            est_str = line_nospaces.split('*')[3]
            act_str = line_nospaces.split('*')[4]

            # cast to int for excel
            est = int(est_str.replace(',', ''))
            act = int(act_str.replace(',', ''))

        if jnum != 'DEFAULT' and desc != 'DEFAULT' and pm != 'DEFAULT' and est != 99999999 and act != 99999999:
            # add job to list and reset variables
            jobs.append(Job(jnum, desc, pm, est, act))

            jnum = 99999999
            desc = 'DEFAULT'
            pm = 'DEFAULT'
            est = 99999999
            act = 99999999

    # reports if there's a discrepancy found so the data can be reviewed
    if len(jobs) != num_jobs:
        print(f'Though {num_jobs} were found, {len(jobs)} were actually reported. You may want to check your data.')

    return jobs


def create_jobs_from_excel_in(data: list[string]):
    orig_job_list_excel: list[Job] = []
    for job in data:
        jnum = int(str(job[0]).replace('-', ''))
        desc = job[2]
        pm = job[3]
        orig_job_list_excel.append(Job(jnum, desc, pm, est=0, act=0))

    return orig_job_list_excel


def compare_jobs(new_jobs: list[Job], old_jobs: list[Job]):  # this assumes both lists are sorted already
    combined_list: list[Job] = []

    same_jobs: bool = False
    nj_greater: bool = False
    oj_greater: bool = False

    for oj in old_jobs:
        for nj in new_jobs:
            # find biggest jnum to append to end of combined list
            if nj.jnum == oj.jnum:
                combined_list.append(Job(oj.jnum, oj.desc, oj.pm, nj.est, nj.act))
            elif nj.jnum > oj.jnum:
                combined_list.append(Job(nj.jnum, nj.desc, nj.pm, nj.est, nj.act))
            else:
                combined_list.append(Job(oj.jnum, oj.desc, oj.pm, oj.est, oj.act))

            break

    return combined_list
