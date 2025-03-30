
# the choices from main redirect here so its a bit easier to comprehend lmao


import file_rw as frw


def add_pdf_data_to_existing_spreadsheet(input_pdf_path, input_excel_path, output_excel_path):
    jobs_from_pdf = frw.read_pdf(input_pdf_path)
    jobs_from_excel = frw.read_excel(input_excel_path)

    frw.create_write_new_excel(jobs_from_pdf, jobs_from_excel, input_excel_path, output_excel_path)


def create_new_excel_from_pdf(input_pdf_path, output_excel_path):
    jobs_from_pdf = frw.read_pdf(input_pdf_path)

    frw.create_write_new_excel(jobs_from_pdf, [], '', output_excel_path)
