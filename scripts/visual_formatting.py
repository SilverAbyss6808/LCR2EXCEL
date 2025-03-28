
# excel sheet formatting stuff here

from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

def format_widths(sheet):
    # hardcoding column widths but can change this later
    sheet.column_dimensions['A'].width = 10.5
    sheet.column_dimensions['B'].width = 10.5
    sheet.column_dimensions['C'].width = 30
    sheet.column_dimensions['D'].width = 8
    sheet.column_dimensions['E'].width = 4
    sheet.column_dimensions['F'].width = 10

    max_col = sheet.max_column

    for col in range(7, max_col + 2):
        sheet.column_dimensions[get_column_letter(col)].width = 10


def color_every_other_line(sheet, color1, color2):
    for i, row in enumerate(sheet.iter_rows()):
        for cell in row:
            if i % 2 == 0:
                cell.fill = PatternFill(patternType="solid", bgColor='00FFFFFF')
            else:
                cell.fill = PatternFill(patternType="solid", bgColor='00123456')