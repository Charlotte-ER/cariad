'''Global variables. Used by both cli and gui.'''

import os
import string

VERSION = '0.0.1'
USER = os.getlogin()

target_directory = None
reference_spreadsheet = None

GUI_TITLE = 'Converting And Renaming Information Access Documents'
GUI_USER_GUIDE = "User Instructions go here"
GUI_ASK_TARGET = "Select target directory"
GUI_ASK_REFERENCE = "Select reference spreadsheet"
PRE_WARNING = "Save and Close Office Applications before running Cariad"

INDEX_SHEET_NAME = 'Index'
INDEX_HEADER_ROW = 3

ALPHABET = string.ascii_lowercase
NUMERALS = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix",
            "x", "xi", "xii", "xiii", "xiv", "xv", "xvi", "xvii",
            "xviii", "xix", "xx", "xxi", "xxii", "xxiii", "xxiv",
            "xxv", "xxvi", "xxvii", "xxviii", "xxix", "xxx"]