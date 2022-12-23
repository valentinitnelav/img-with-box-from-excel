# Description: This file contains the GUI for the Boxcel application.
# The GUI consists of a single window with a button which opens a file dialog that 
# allows the user to select an Excel file.
# The file path to the Excel file is passed to the start_project function that is 
# defined in the start_project.py file.
# The start_project function is responsible for generating the Python code that
# enables the xlwings functionality for the Excel file.


import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox # see issue #17
from start_project import *


# Create root window
root_window = tk.Tk()
# Root window title and dimension
root_window.title("Boxcel")
# Set geometry(width x height)
root_window.geometry('600x300')

# Add an informative label to the root window
lbl_info = tk.Label(root_window, text = "Choose an Excel file to open:")
lbl_info.pack(side=tk.TOP, anchor='w') 
# anchor='w' means left align (west)

# Button widget to browse for the Excel file
btn = tk.Button(root_window, text="Browse & execute")
btn.pack(side=tk.TOP, anchor='w')

# Label to announce the selected file. 
# I gets updated with text when a file is selected.
lbl_file = tk.Label(root_window)
lbl_file.pack(side=tk.TOP, anchor='w')

# Message widget to display the selected file path
msg_file_path = tk.Message(root_window)
msg_file_path.pack(side=tk.TOP, anchor='w')


# Function to open a file dialog and select an Excel file.
# The file path to the Excel file is passed to the start_project function.
def open_file():
    file = filedialog.askopenfilename(
        parent=root_window, 
        initialdir='C:\\', # On Linux shows the current directory
        title='Choose a file',
        filetypes=[('Excel files', '*.xlsx')]
    )
    # Print for diagnostic purposes
    print('Selected file:', file)

    # If no file was selected, then file contains an empty tuple. Test for this 
    # to avoid errors.
    if len(file) != 0:
        # Update the labels and message widgets
        lbl_file.config(text = "Selected file:")
        msg_file_path.config(bg='gray80', width=600, text = file)

        # Call the start_project function from start_project.py
        # and pass the file path to the Excel file as an argument.
        # If the start_project function was successful, print that the operation is done,
        # otherwise print the error message.
        try:
            start_project(file)
            # Update the label to announce the result of the operation
            tk.messagebox.showinfo(
                title = "Information", 
                message = "All good! Python code generated. Choose another file or close the application."
                )
        except Exception as e:
            tk.messagebox.showerror(title = "Error!", message = str(e))

# Execute the open_file function when the button is clicked
btn.config(command=open_file)

# Execute Tkinter
root_window.mainloop()