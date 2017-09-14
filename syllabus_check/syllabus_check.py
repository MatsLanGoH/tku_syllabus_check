# encoding: utf-8
import os

# Excel workbook handling
from xlrd import open_workbook

# Word document handling
from docx import Document
from docx.shared import Inches, RGBColor
from docx.text.run import Font, Run

# Data positions in excel workbooks
from positions import *

# Check instructions
from instructions import *
from check import Check

# Constants
RESULTS_FOLDER = u"results"


def check_sheet(sheet):
    """
    Checks a single worksheet.
    """
    # TODO: Implement this.

    # Step 1: Setup environment
    # (Setup positions and plan grid according to specs)

    # Step 2: Execute checks in checks.
    # NOTE: Make sure we don't have to recreate the check_series each time.
    pass


def execute_checks(workbook):
    """
    Executes checks for a given workbook.
    """
    print(workbook)
    
    # Open workbook
    xl_workbook = open_workbook(workbook)
    result_name = workbook.split('.')[0] + '-tenken.docx'
    result_location = os.path.join('.', RESULTS_FOLDER, result_name)

    # Template file for result output
    # TODO: NOTE: Make sure this file exists in our path
    # result_doc = Document(os.path.join(os.environ['RESOURCEPATH'], 'templates/default.docx'))
    result_doc = Document()
    result_doc.add_heading('シラバス点検 自動点検結果', 0)

    # Get the sheets from the workbook
    sheets = xl_workbook.sheets()

    # Execute checks for each sheet
    for sheet in sheets:
        # TODO: Implement this
        check_sheet(sheet)


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
        try:
            execute_checks(wb)
        except Exception as e:
            print("Error:", e)


if __name__ == "__main__":
    main()
