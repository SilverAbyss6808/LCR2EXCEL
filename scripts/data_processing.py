# this is where the stuff for processing the input pdf is gonna go
# plus any other data processing things

import openpyxl as opxl
import string


alphabet: dict = {
    'A': 1,
    'B': 2,
    'C': 3,
    'D': 4,
    'E': 5,
    'F': 6,
    'G': 7,
    'H': 8,
    'I': 9,
    'J': 10,
    'K': 11,
    'L': 12,
    'M': 13,
    'N': 14,
    'O': 15,
    'P': 16,
    'Q': 17,
    'R': 18,
    'S': 19,
    'T': 20,
    'U': 21,
    'V': 22,
    'W': 23,
    'X': 24,
    'Y': 25,
    'Z': 26
}


class Job:
    jnum: int  # will need to cast to string with - for excel sheet (XX-XX-XXXX)
    desc: string
    pm: string
    est: int
    act: int
    prev_costs: dict = {}

    def __init__(self, jnum: int, desc: string, pm: string, est: int, act: int, prev_costs: dict = None):
        self.jnum = jnum
        self.desc = desc
        self.pm = pm
        self.est = est
        self.act = act
        self.prev_costs = prev_costs

    def __str__(self):
        return (f'Job Number: {self.jnum}, Project Manager: {self.pm}, Description: {self.desc}\n'
                f'Estimated Cost: {self.est}, Actual Cost: {self.act}\n'
                f'Previous Costs: {self.prev_costs}\n')


class JobRow:
    column1: string = ''  # jnum
    jobno: string = ''  # should be the value of column1 every time
    desc: string = ''  # name of job
    column2: string = ''  # pm
    pm: string = ''  # should be the value of column2 every time
    column5: string = ''  # 0=Estimate, 1=Actual, 2=Last Week, 3=Remaining
    prev_costs: dict = {}
    current_est: int = 0
    current_act: int = 0

    def __init__(self, column1, jobno, desc, column2, pm, column5, prev_costs, current_est, current_act):
        self.column1 = column1
        self.jobno = jobno
        self.desc = desc
        self.column2 = column2
        self.pm = pm
        self.column5 = column5
        self.prev_costs = prev_costs
        self.current_est = current_est
        self.current_act = current_act


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
            jobs.append(Job(jnum, desc, pm, est, act, prev_costs=None))  # no previous costs cause new data

            jnum = 99999999
            desc = 'DEFAULT'
            pm = 'DEFAULT'
            est = 99999999
            act = 99999999

    # reports if there's a discrepancy found so the data can be reviewed
    if len(jobs) != num_jobs:
        print(f'Though {num_jobs} were found, {len(jobs)} were actually reported. You may want to check your data.')

    return jobs


def create_jobs_from_excel_in(data: list[string], max_col: int):
    orig_job_list_excel: list[Job] = []
    est, act = 0, 0

    for index, job in enumerate(data):
        idx_mod = index % 4
        if idx_mod == 0:
            jnum = None
            if job[0] is not None:
                jnum = int(str(job[0]).replace('-', ''))

            desc = job[2]
            pm = job[3]

            prev_costs: dict = {}
            prev_est: dict = {}
            prev_job_estimate = -1

            for inner_idx, col in enumerate(range(6, max_col)):  # estimated costs start at 6 on line 0
                if job[col] == prev_job_estimate:
                    prev_est[inner_idx] = f'prev{inner_idx}'
                else:
                    prev_est[inner_idx] = job[col]
                    prev_job_estimate = job[col]

        elif idx_mod == 1:
            for inner_idx, col in enumerate(range(6, max_col)):  # actual costs start at 6 on line 1
                prev_costs[prev_est[inner_idx]] = job[col]

            orig_job_list_excel.append(Job(jnum, desc, pm, est, act, prev_costs))

    return orig_job_list_excel


def compare_jobs(new_jobs: list[Job], old_jobs: list[Job]):  # this assumes both lists are sorted already
    combined_list: list[Job] = []

    old_idx = 0
    old_max = len(old_jobs)

    new_idx = 0
    new_max = len(new_jobs)

    while old_idx <= old_max - 1 and new_idx <= new_max - 1:
        if old_jobs[old_idx].jnum < new_jobs[new_idx].jnum:
            combined_list.append(old_jobs[old_idx])
            old_idx += 1
        else:
            combined_list.append(new_jobs[new_idx])
            new_idx += 1
            if old_jobs[old_idx].jnum == new_jobs[new_idx - 1].jnum:
                old_idx += 1

    return combined_list


def format_jobs_as_excel(list_to_format: list[Job]):
    formatted_job_list: list = []

    for index, job in enumerate(list_to_format, 2):  # starting at 2 to leave 1 for title
        row_mod = index % 4

        match row_mod:
            case 2:  # first row of four, should have all info but actual costs
                jnum_formatted = str(job.jnum)
                jnum_formatted = jnum_formatted[:2] + '-' + jnum_formatted[2:4] + '-' + jnum_formatted[4:]

                row = [jnum_formatted, f'=A{index}', job.desc, job.pm, f'=D{index}', 'Estimate']

                if job.prev_costs is not None:
                    for est in job.prev_costs.keys():
                        row.append(est)
            case 3:  # second row, should have actual costs and formulas
                row = ['', f'=A{index - 1}', '', '', f'=D{index - 1}', 'Actual']

                if job.prev_costs is not None:
                    for act in job.prev_costs.values():
                        row.append(act)
            case 0:  # third row, should be almost all formulas
                row = ['', f'=A{index - 2}', '', '', f'=D{index - 2}', 'Last Week', 0]

                first_cell_letter = 'H'
                second_cell_letter = 'G'

                if job.prev_costs is not None:
                    for ind in range(1, len(job.prev_costs)):
                        first_key = (alphabet.get('H') + (ind + 1)) % 26
                        second_key = (alphabet.get('G') + ind) % 26

                        first_cell_letter = [key for key, val in alphabet.items() if val == first_key]
                        second_cell_letter = [key for key, val in alphabet.items() if val == second_key]

                        first_cell_letter = str(first_cell_letter).replace('[\'', '').replace('\']', '')
                        second_cell_letter = str(second_cell_letter).replace('[\'', '').replace('\']', '')

                        row.append(f'={first_cell_letter}{index - 1}-{second_cell_letter}{index - 1}')

            case 1:  # fourth row, should be almost all formulas
                row = ['', f'=A{index - 3}', '', '', f'=D{index - 3}', 'Remaining']

                if job.prev_costs is not None:
                    for ind in range(0, len(job.prev_costs)):
                        new_key = (alphabet.get('G') + ind) % 26
                        cell_letter = [key for key, val in alphabet.items() if val == new_key]
                        cell_letter = str(cell_letter).replace('[\'', '').replace('\']', '')
                        row.append(f'={cell_letter}{index - 3}-{cell_letter}{index - 2}')

            case _:  # default case. inform user of error?
                row = ''

        formatted_job_list.append(row)

    return formatted_job_list
