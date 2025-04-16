# this is where the stuff for processing the input pdf is gonna go
# plus any other data processing things

import string


pdf_date: string = ''
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
    prev_ests: list[int] = []
    prev_acts: list[int] = []
    note: list[(int, str)]
    grouped: bool = False
    hidden: bool = False

    def __init__(self, jnum: int, desc: string, pm: string, est: int, act: int, prev_ests: list[int],
                 prev_acts: list[int], note: list[(int, str)], grouped: bool, hidden: bool):
        self.jnum = jnum
        self.desc = desc
        self.pm = pm
        self.est = est
        self.act = act
        self.prev_ests = prev_ests
        self.prev_acts = prev_acts
        self.note = note
        self.grouped = grouped
        self.hidden = hidden

    def __str__(self):
        return (f'Job Number: {self.jnum}, Project Manager: {self.pm}, Description: {self.desc}\n'
                f'Estimated Cost: {self.est}, Actual Cost: {self.act}\n'
                f'Previous Estimates: {self.prev_ests}\n'
                f'Previous Actuals: {self.prev_acts}\n'
                f'Note(s): {self.note}\n'
                f'Is grouped: {self.grouped}. Is hidden: {self.hidden}.\n')


class JobRow:
    column1: string = ''  # jnum
    jobno: string = ''  # should be the value of column1 every time
    desc: string = ''  # name of job
    column2: string = ''  # pm
    pm: string = ''  # should be the value of column2 every time
    column5: string = ''  # 0=Estimate, 1=Actual, 2=Last Week, 3=Remaining
    prev_ests: list[int] = []
    prev_acts: list[int] = []
    current_est: int = 0
    current_act: int = 0

    def __init__(self, column1, jobno, desc, column2, pm, column5, prev_ests: list[int],
                 prev_acts: list[int], current_est, current_act):
        self.column1 = column1
        self.jobno = jobno
        self.desc = desc
        self.column2 = column2
        self.pm = pm
        self.column5 = column5
        self.prev_ests = prev_ests
        self.prev_acts = prev_acts
        self.current_est = current_est
        self.current_act = current_act


def format_pdf_data_as_job(data: string):  # return an array of jobs
    num_jobs: int = 0
    data_lines = data.split('\n')
    date_found = False
    global pdf_date

    for index, line in enumerate(data_lines):
        if not date_found and 'System Date:' in line:
            date_found = True
            pdf_date = line.split(' ')[2].replace('-', '/')

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
            try:
                jnum = int(jnum_string.replace('-', ''))
            except ValueError:
                cc = False
                continue

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
            jobs.append(Job(jnum, desc, pm, est, act, prev_ests=[], prev_acts=[], note=[None, None], grouped=False, hidden=False))

            jnum = 99999999
            desc = 'DEFAULT'
            pm = 'DEFAULT'
            est = 99999999
            act = 99999999

    # reports if there's a discrepancy found so the data can be reviewed
    if len(jobs) != num_jobs:
        # print(f'Though {num_jobs} were found, {len(jobs)} were actually reported. You may want to check your data.')
        pass
    return jobs


def create_jobs_from_excel_in(data: list[string], max_col: int):
    orig_job_list_excel: list[Job] = []
    est, act = 0, 0
    note: list[(int, str)] = [None, None]
    grouped = False
    hidden = False

    for index, job in enumerate(data):
        idx_mod = index % 4  # finds what row of the job it is

        # note stuff and group states
        if idx_mod == 0:  # reset list of notes/state of group at beginning of job
            note = [None, None]
            grouped = False
            hidden = False
        if job[max_col - 1] is not None:
            note.append((idx_mod, job[max_col - 1]))

        # adds data to job according to what job row it is
        if idx_mod == 0:  # first row of job. has most of the info
            jnum = None
            if job[0] is not None:
                jnum = int(str(job[0]).replace('-', ''))

            desc = job[2]
            pm = job[3]

            prev_est: list[int] = []
            prev_act: list[int] = []

            lastint = 0
            for col in range(6, max_col):  # estimated costs start at 6 on first row of four
                if isinstance(job[col], int) and job[col] != 0:
                    prev_est.append(job[col])
                    lastint = job[col]
                elif col == max_col - 1:  # this is the notes row, should break to allow note to be copied
                    break
                elif isinstance(job[col], str) or job[col] == 0 or job[col] is None:
                    prev_est.append(lastint)
                else:
                    prev_est.append(0)

            # this takes care of grouping and hiding previously grouped/hidden jobs
            if 'g' in job[max_col]:
                grouped = True
            if 'h' in job[max_col]:
                hidden = True

        elif idx_mod == 1:  # second row of job. only has actual costs (and maybe notes but thats not processed here)
            lastint = 0
            for col in range(6, max_col):  # actual costs start at 6 on second row of four
                if isinstance(job[col], int) and job[col] != 0:
                    prev_act.append(job[col])
                    lastint = job[col]
                elif col == max_col - 1:  # this is the notes row, should break to allow note to be copied
                    break
                elif job[col] == 0 or job[col] is None:
                    prev_act.append(lastint)
                else:
                    prev_act.append(0)

        elif idx_mod == 3:  # this is the last row of a job, so all info will be known
            orig_job_list_excel.append(Job(jnum, desc, pm, est, act, prev_est, prev_act, note, grouped, hidden))
            print(str(Job(jnum, desc, pm, est, act, prev_est, prev_act, note, grouped, hidden)))

        else:  # only the third line. no new info here
            pass

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
        elif old_jobs[old_idx].jnum > new_jobs[new_idx].jnum:
            combined_list.append(new_jobs[new_idx])
            new_idx += 1
        else:
            combined_list.append(Job(new_jobs[new_idx].jnum, new_jobs[new_idx].desc, new_jobs[new_idx].pm, new_jobs[new_idx].est,
                                     new_jobs[new_idx].act, old_jobs[old_idx].prev_ests, old_jobs[old_idx].prev_acts,
                                     old_jobs[old_idx].note, old_jobs[old_idx].grouped, old_jobs[old_idx].hidden))
            new_idx += 1
            old_idx += 1

    return combined_list


def format_jobs_as_excel(list_to_format: list[Job], max_col: int):
    formatted_job_list: list = []
    index = 2  # starting at row 2 to leave 1 for title

    for job in list_to_format:
        for i in range(0, 4):
            match i:
                case 0:  # first row of four, should have all info but actual costs

                    jnum_formatted = str(job.jnum)
                    jnum_formatted = jnum_formatted[:2] + '-' + jnum_formatted[2:4] + '-' + jnum_formatted[4:]

                    row = [jnum_formatted, f'=A{index}', job.desc, job.pm, f'=D{index}', 'Estimate']

                    col_filled = 6
                    current_prev = 0

                    for est in job.prev_ests:
                        row.append(est)
                        current_prev = est
                        col_filled += 1

                    # really shouldnt execute :P leaving it in just in case though
                    while (max_col - col_filled) > 0:
                        row.append('')
                        col_filled += 1

                    if job.est != 0:
                        row.append(job.est)
                    else:
                        row.append(current_prev)

                case 1:  # second row, should have actual costs and formulas
                    row = ['', f'=A{index - 1}', '', '', f'=D{index - 1}', 'Actual']

                    col_filled = 6
                    current_prev = 0

                    for act in job.prev_acts:
                        row.append(act)
                        col_filled += 1
                        current_prev = act

                    while (max_col - col_filled) > 0:
                        row.append('')
                        col_filled += 1

                    if job.act != 0:
                        row.append(job.act)
                    else:
                        row.append(current_prev)

                case 2:  # third row, should be almost all formulas
                    row = ['', f'=A{index - 2}', '', '', f'=D{index - 2}', 'Last Week']

                    if max_col != 0:
                        row.append(0)
                        for ind in range(6, max_col):
                            first_key = (alphabet.get('H') + (ind - 6)) % 26
                            second_key = (alphabet.get('G') + (ind - 6)) % 26

                            first_cell_letter = [key for key, val in alphabet.items() if val == first_key]
                            second_cell_letter = [key for key, val in alphabet.items() if val == second_key]

                            first_cell_letter = str(first_cell_letter).replace('[\'', '').replace('\']', '')
                            second_cell_letter = str(second_cell_letter).replace('[\'', '').replace('\']', '')

                            row.append(f'=IF(AND(ISNUMBER({first_cell_letter}{index-1}),ISNUMBER({second_cell_letter}{index-1})), '
                                       f'{first_cell_letter}{index-1}-{second_cell_letter}{index-1}, 0)')

                case 3:  # fourth row, should be almost all formulas
                    row = ['', f'=A{index - 3}', '', '', f'=D{index - 3}', 'Remaining']

                    for ind in range(7, max_col + 2):  # +2 because 1 for new column and 1 to make sure it actually runs
                        new_key = (alphabet.get('G') + (ind - 7)) % 26
                        cell_letter = [key for key, val in alphabet.items() if val == new_key]
                        cell_letter = str(cell_letter).replace('[\'', '').replace('\']', '')
                        row.append(f'={cell_letter}{index - 3}-{cell_letter}{index - 2}')

                case _:  # default case. it iterates through 0, 1, 2, and 3 so its literally not possible to get here
                    row = ''

            for nt in job.note:
                if nt is not None and nt[0] == i:  # skip defaults
                    # print(nt[1])
                    row.append(nt[1])

            index += 1
            formatted_job_list.append(row)

    return formatted_job_list
