'''Execute the rename and convert process. Used by both cli and gui.'''
import os
import re
import pandas as pd
import sys

from openpyxl import load_workbook
from tkinter import messagebox

import variables

def run_cariad(target_directory, reference_spreadsheet, mode):
    index = get_index_dataframe_from_spreadsheet(reference_spreadsheet, mode)


def tell_user(message, mode):
    if mode == 'cli':
        print(message)
    else:
        messagebox.showwarning(title = "Message",
                        message = message)


def get_index_dataframe_from_spreadsheet(reference_spreadsheet, mode):
    _, ext = os.path.splitext(reference_spreadsheet)
    matches = re.search(r'.xlsx?$', ext.lower())
    if not matches:
        tell_user('Not a spreadsheet.', mode)
        sys.exit(1)

    workbook = load_workbook(filename=reference_spreadsheet)

    try:
        sheet = workbook[variables.INDEX_SHEET_NAME]
    except:
        tell_user(f'Spreadsheet has no tab called: {variables.INDEX_SHEET_NAME}', mode)
        sys.exit(1)

    index = pd.DataFrame(get_index_rows_from_spreadsheet(sheet, variables.INDEX_HEADER_ROW))
    index.columns = index.iloc[0]
    index = index[1:]

    return index


def get_index_rows_from_spreadsheet(sheet, header):
    for count, row in enumerate(sheet.values, start=1):
        if count < header:
            continue
        yield row

