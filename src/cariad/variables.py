'''Global variables. Used by both cli and gui.'''

import os

VERSION = '0.0.1'
USER = os.getlogin()

target_directory = None
reference_spreadsheet = None

GUI_TITLE = 'Converting And Renaming Information Access Documents'
GUI_USER_GUIDE = "User Instructions go here"
GUI_ASK_TARGET = "Select target directory"
GUI_ASK_REFERENCE = "Select reference spreadsheet"

INDEX_SHEET_NAME = 'Index'
INDEX_HEADER_ROW = 3