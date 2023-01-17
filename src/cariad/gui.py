'''Launch cariad from gui instead of cli.'''

from tkinter import *
from tkinter import filedialog, messagebox, ttk

import run
import variables


def run_gui():
    root = Tk()
    root.geometry('900x700+100+100')
    root.title('cariad')

    title_frame = ttk.Frame(root)
    title_frame.pack()
    title_frame.config(height = 100, 
                        width = 900,
                        padding = (30,15))

    ttk.Label(title_frame, text = variables.GUI_TITLE,
                    wraplength = 850,
                    justify = CENTER,
                    font = ('Calibri', 12, 'bold')).pack()

    ttk.Label(title_frame, text = variables.VERSION,
                    wraplength = 850,
                    justify = CENTER,
                    font = ('Calibri', 10, 'italic')).pack()
    
    ttk.Label(title_frame, text = variables.GUI_USER_GUIDE,
                    wraplength = 600,
                    justify = CENTER,
                    font = ('Calibri', 10)).pack()

    ask_target_directory_message = ttk.Label(title_frame, 
                                    text = variables.GUI_ASK_TARGET,
                                    wraplength = 500,
                                    justify = LEFT,
                                    font = ('Calibri', 10))
    ask_target_directory_message.pack()

    ask_target_directory_button = ttk.Button(title_frame,
                                    text = "Browse for target directory",
                                    command = get_target_directory,
                                    state = NORMAL)
    ask_target_directory_button.pack(padx = 10)

    ask_reference_spreadsheet_message = ttk.Label(title_frame, 
                                    text = variables.GUI_ASK_REFERENCE,
                                    wraplength = 500,
                                    justify = LEFT,
                                    font = ('Calibri', 10))
    ask_reference_spreadsheet_message.pack()

    ask_reference_spreadsheet_button = ttk.Button(title_frame,
                                    text = "Browse for reference spreadsheet",
                                    command = get_reference_spreadsheet,
                                    state = NORMAL)
    ask_reference_spreadsheet_button.pack(padx = 10)

    run_button = ttk.Button(title_frame, 
                            text = "Run Cariad",
                            command = run_command,
                            state = NORMAL)
    run_button.pack(padx = 10, pady = 10)

    root.mainloop()


def get_target_directory():
    default = f'C:/Users/{variables.USER}/Documents'
    variables.target_directory = filedialog.askdirectory(initialdir = default)
    return variables.target_directory


def get_reference_spreadsheet():
    variables.reference = filedialog.askopenfilename(initialdir = variables.target_directory)
    return variables.reference


def run_command():
    run.run_cariad(variables.target_directory, variables.reference_spreadsheet, 'gui')
    return True


if __name__ == '__main__':
    run_gui()