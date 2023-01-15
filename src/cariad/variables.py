'''Global variables. Used by both cli and gui.'''

import os

# App Info
VERSION = '0.0.1'

# Runtime variables
USER = os.getlogin()
target_directory = None
reference_spreadsheet = None