'''Launch cariad from gui instead of cli.'''

import tkinter
from tkinter import filedialog, Tk, ttk

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
                    justify = tkinter.CENTER,
                    font = ('Calibri', 12, 'bold')).pack()

    ttk.Label(title_frame, text = variables.VERSION,
                    wraplength = 850,
                    justify = tkinter.CENTER,
                    font = ('Calibri', 10, 'italic')).pack()
    
    ttk.Label(title_frame, text = variables.GUI_USER_GUIDE,
                    wraplength = 600,
                    justify = tkinter.CENTER,
                    font = ('Calibri', 10)).pack()

    ask_target_directory_message = ttk.Label(title_frame, 
                                    text = variables.GUI_ASK_TARGET,
                                    wraplength = 500,
                                    justify = tkinter.LEFT,
                                    font = ('Calibri', 10))
    ask_target_directory_message.pack()

    ask_target_directory_button = ttk.Button(title_frame,
                                    text = "Browse for target directory",
                                    command = get_target_directory,
                                    state = tkinter.NORMAL)
    ask_target_directory_button.pack(padx = 10)

    ask_reference_spreadsheet_message = ttk.Label(title_frame, 
                                    text = variables.GUI_ASK_REFERENCE,
                                    wraplength = 500,
                                    justify = tkinter.LEFT,
                                    font = ('Calibri', 10))
    ask_reference_spreadsheet_message.pack()

    ask_reference_spreadsheet_button = ttk.Button(title_frame,
                                    text = "Browse for reference spreadsheet",
                                    command = get_reference_spreadsheet,
                                    state = tkinter.NORMAL)
    ask_reference_spreadsheet_button.pack(padx = 10)

    ttk.Label(title_frame, text = variables.PRE_WARNING,
                    wraplength = 600,
                    justify = tkinter.CENTER,
                    font = ('Calibri', 10)).pack()

    run_button = ttk.Button(title_frame, 
                            text = "Run Cariad",
                            command = run_command,
                            state = tkinter.NORMAL)
    run_button.pack(padx = 10, pady = 10)

    root.mainloop()


def get_target_directory():
    default = f'C:/Users/{variables.USER}/Documents'
    variables.target_directory = filedialog.askdirectory(initialdir = default)
    return variables.target_directory


def get_reference_spreadsheet():
    variables.reference_spreadsheet = filedialog.askopenfilename(initialdir = variables.target_directory)
    return variables.reference_spreadsheet


def run_command():
    run.run_cariad(variables.target_directory, variables.reference_spreadsheet, 'gui')
    return True


if __name__ == '__main__':
    run_gui()