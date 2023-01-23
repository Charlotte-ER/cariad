'''Execute the rename and convert process. Used by both cli and gui.'''

from tkinter import messagebox

def run_cariad(target_directory, reference_spreadsheet, mode):
    ...
    
    # Handle incorrect usage

    # Rename documents

    # Get email attachments

    # Convert to pdf


def tell_user(message, mode):
    if mode == 'cli':
        print(message)
    else:
        messagebox.showwarning(title = "Message",
                        message = message)


if __name__ == '__main__':
    run_cariad('boo', 'faa', 'cli')
