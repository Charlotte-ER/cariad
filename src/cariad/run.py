'''Execute the rename and convert process. Used by both cli and gui.'''
import os
import re
import pandas as pd
import sys
import win32com.client

from openpyxl import load_workbook
from tkinter import messagebox

import variables

def run_cariad(target_directory, reference_spreadsheet, mode): 
    os.chdir(target_directory)
    index = get_index_dataframe_from_spreadsheet(reference_spreadsheet, mode)

    # For each row in the index
    for i in range(len(index)):
        print_num = index.iloc[i]['Print Index Number']
        doc_num = index.iloc[i]['Doc. Number']
        version = index.iloc[i]['Version']
        extension = index.iloc[i]['Extension'].lower()

        old_name = f'{target_directory}/{doc_num}_{version}.{extension}'
        new_name = f'{target_directory}/{print_num}.{extension}'

        if os.path.isfile(old_name):
            os.rename(old_name, new_name)
        
        if extension == 'msg':
            convert_email(new_name)
    
    convert_to_pdf()


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


def convert_email(email, level=1):
    printnum, _ = os.path.splitext(email)

    if level % 2 != 0:
        label = variables.ALPHABET
    else:
        label = variables.NUMERALS
    
    counter = 0

    outlook = win32com.client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
    message = outlook.OpenSharedItem(email)

    for attachment in message.Attachments:
        _, ext = os.path.splitext(attachment.FileName)

        name = f'{printnum}({label[counter]}){ext}'
        attachment.SaveAsFile(name)
        counter += 1

        if ext.lower() == ".msg":
            convert_email(name, level + 1)

    message = None
    outlook = None

def convert_to_pdf():
    # Goes through directory and converts supported files to pdf
    for file in os.listdir():
        filepath = os.path.abspath(file)
        root, ext = os.path.splitext(file)

        # Identify PowerPoint formats and convert to pdf
        if ext.lower() in [".pptx", ".potx", "ppsx", ".thmx", ".ppt", ".pot", ".pps"]:
            try:
                ppt_to_pdf(filepath)
            except Exception as error:
                continue

        # Identify Word formats and convert to pdf
        if ext.lower() in [".doc", ".dot", ".wbk", ".docx", ".docm", ".dotx",
                ".dotm", ".docb", ".DOCX"]:
            try:
                word_to_pdf(filepath)
            except Exception as error:
                continue

def ppt_to_pdf(filename):
    # Converts PowerPoint document to PDF
    root, ext = os.path.splitext(filename)
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    document = powerpoint.Presentations.Open(filename, ReadOnly=False)
    document.SaveAs(f'{root}.pdf', 32)
    document.Close()
    powerpoint.Quit()
    document = None
    powerpoint = None
    os.remove(filename)


def word_to_pdf(filename):
    # Converts Word document to PDF
    word = win32com.client.Dispatch("Word.Application")
    # word.Visible = False  '''Leaving False until can handle pop ups'''
    document = word.Documents.Open(filename, ReadOnly=False)
    root, ext = os.path.splitext(filename)
    document.SaveAs(f'{root}.pdf', 17)
    document.Close(SaveChanges=False)
    word.Quit()
    document = None
    word = None
    os.remove(filename)