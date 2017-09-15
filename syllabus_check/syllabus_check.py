# encoding: utf-8
import os

# Excel workbook handling
from xlrd import open_workbook

# Word document handling
from docx import Document
from docx.shared import Inches, RGBColor
from docx.text.run import Font, Run

# Data positions in excel workbooks
from positions import kamoku_info, kamoku_maps

# Check instructions
from instructions import checks
from check import Check

# Constants
RESULTS_FOLDER = u"results"


def check_sheet(sheet):
    """
    Checks a single worksheet.
    """
    # Step 1: Setup environment
    k_items = kamoku_maps[sheet.nrows] 

    result = ""
    # Step 2: Execute checks in checks.
    # NOTE: Make sure we don't have to recreate the check_series each time.
    for check in checks:
        ck = Check(*check, k_items, sheet)
        if not ck.passes():
            result += '\n□ ' + ck.msg
            result += '\n' + ck.errors

    print(result)
    return result


def execute_checks(workbook):
    """
    Executes checks for a given workbook.
    Saves result to a docx file containing the name of the source workbook
    as its file name.
    """
    # Open workbook and get its name to create output file name
    xl_workbook = open_workbook(workbook)
    result_name = workbook.split('.')[0] + '-tenken.docx'
    result_location = os.path.join('.', RESULTS_FOLDER, result_name)

    # Setup output file
    document = Document()
    document.add_heading('シラバス点検 自動点検結果', 0)

    # Get the sheets from the workbook
    sheets = xl_workbook.sheets()

    # Execute checks for each sheet
    for sheet in sheets:
        # Build strings for result output
        teacher_str = u'\n\t'.join(sheet.cell_value(*kamoku_info['teacher'])
                                   .split('\n'))
        class_str = sheet.cell_value(*kamoku_info['name'])

        # Get check results
        result = check_sheet(sheet)

        # Add results to document
        document.add_heading('{}: {}'.format(sheet.name, teacher_str), level=2)
        p = document.add_paragraph()
        p.add_run('授業科目名: {}'.format(class_str)).bold = True
        p.paragraph_format.left_indent = Inches(0.5)
        p.add_run('{}'.format(result))

        # Format colors according to result status
        if len(result) == 0:
            p.add_run('\n => 点検OK。')
        else:
            p.add_run('\n => 上記の項目を確認し、必要な場合は修正を行ってください。')\
                    .font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        p.add_run('\n' + '_' * 78 + '\n')  # Break line

    # Finally, save result file
    document.save(result_location)


def main():
    """
    Finds all Workbooks in directory and checks the syllabus.
    """
    # TODO: Allow the user to specify a path.
    # Read workbooks in directory
    workbooks = [f for f in os.listdir('.')
                 if os.path.isfile(f)
                 and f.lower().endswith('.xls')]

    # TODO: Do we want to escape like this?
    if len(workbooks) == 0:
        print('No workbooks in directory.')
        return

    # Create output directory if necessary
    if not os.path.exists(RESULTS_FOLDER) and len(workbooks) > 0:
        os.makedirs(os.path.join(RESULTS_FOLDER))

    # Check syllabus for all workbooks.
    for wb in workbooks:
        # try:
        execute_checks(wb)
        # except Exception as e:
        #     print("Error:", e)


if __name__ == "__main__":
    main()
