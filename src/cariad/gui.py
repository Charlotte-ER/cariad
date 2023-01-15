'''Run cariad from gui instead of cli.'''
import os
from tkinter import *
from tkinter import filedialog, messagebox, ttk

import globals

def run_gui():
    '''Set up Root Window'''
    root = Tk()
    root.geometry('900x700+100+100')
    root.title('cariad')


    '''Set up Title Frame and Text Labels'''
    # Title Frame
    title_frame = ttk.Frame(root)
    title_frame.pack()
    title_frame.config(height = 100, 
                        width = 900,
                        padding = (30,15))

    # Title Details
    ttk.Label(title_frame, text = "Converting And Renaming Information Access Documents",
                    wraplength = 850,
                    justify = CENTER,
                    font = ('Calibri', 12, 'bold')).pack()

    # Version Details
    ttk.Label(title_frame, text = 'link to global version goes here',
                    wraplength = 850,
                    justify = CENTER,
                    font = ('Calibri', 10, 'italic')).pack()
    
    # Introduction
    ttk.Label(title_frame, text = "User Instructions go here",
                    wraplength = 600,
                    justify = CENTER,
                    font = ('Calibri', 10)).pack()

    # Label to ask User for Case Folder
    selectFolderLabel = ttk.Label(title_frame, 
                                    text = "Select target directory",
                                    wraplength = 500,
                                    justify = LEFT,
                                    font = ('Calibri', 10))
    selectFolderLabel.pack()

    # Browse Button
    getFolderButton = ttk.Button(title_frame,
                                    text = "Browse",
                                    command = get_target_directory,
                                    state = NORMAL)
    getFolderButton.pack(padx = 10)

    root.mainloop()




def get_target_directory():
    target = filedialog.askdirectory(initialdir = f'C:/Users/{os.getlogin()}/Documents')
    globals.target_directory = target
    return target


if __name__ == '__main__':
    run_gui()