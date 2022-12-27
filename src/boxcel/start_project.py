# This is the main script that is executed when the user clicks the button 
# "Browse & execute" in the GUI.
# The start_project function from this script is responsible for generating the 
# Python code that enables the xlwings functionality for the Excel file.
# This script can also be executed from the command line like this:
# python3 start_project.py path/to/file.xlsx
# or
# python start_project.py path/to/file.xlsx

import sys
import os

# Function to get the absolute path to the current working directory
# This is to fix errors like this:
# "[Error 2] No such file or directory:
# 'C:\\Users\\user\\AppData\\Local\\Temp\\_MEI112962\\base_library.zip\\display_images.py'"
# This is because the script is executed from a temporary directory when using PyInstaller.
# PyInstaller creates a temp folder and stores path in _MEIPASS2 or ._MEIPASS
# See https://stackoverflow.com/a/13790741/5193830
# Also https://pyinstaller.org/en/stable/runtime-information.html#run-time-information
def get_work_dir_path():
    """ Get absolute path to current working directory"""
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = sys._MEIPASS
    else:
        # When running in "normal" Python environment.
        # For example, when running from the command line:
        # python3 start_project.py path/to/file.xlsx
        # sys.path[0] returns something like path/to/src/boxcel
        base_path = sys.path[0]

    return base_path


def start_project(xlsx_file):
    """ Writes the needed code to the *.py file corresponding to the *.xlsx file.
    The code is copied from the script display_images.py and the name of the
    corresponding xlsx file is inserted after the line if __name__ == "__main__":
    """

    # Test if the xlsx file exists
    if not os.path.isfile(xlsx_file):
        print("The file " + xlsx_file + " does not exist or wrong path.")
        sys.exit()

    # Test if the xlsx file has the extension .xlsx
    if not xlsx_file.endswith(".xlsx"):
        print("The file " + xlsx_file + " is not an xlsx file.")
        sys.exit()

    # Read the path to the folder where the xlsx file is located
    path_to_xlsx_file = os.path.dirname(xlsx_file)

    # Read the name of the xlsx file
    xlsx_file_name = os.path.basename(xlsx_file)

    # Read the name of the xlsx file without the extension
    xlsx_file_name_without_extension = os.path.splitext(xlsx_file_name)[0]


    # Generate the needed Python file corresponding to the *.xlsx file.
    # So, copy the code from the script display_images.py and 
    # insert the name of the corresponding xlsx file after the line
    # if __name__ == "__main__": 
    
    # Get the path to the directory where this script file is located & executed from:
    print('get_work_dir_path():', get_work_dir_path()) # print kept for debugging purposes
    display_images_py_file = os.path.join(get_work_dir_path(), "display_images.py")

    target_py_file = os.path.join(path_to_xlsx_file, xlsx_file_name_without_extension + ".py")

    with open(display_images_py_file,'r') as firstfile, open(target_py_file,'w') as secondfile:
        # Read content from first file
        for line in firstfile:  
            # For the particular line containing "xw.Book("xlwings_test_project.xlsm").set_mock_caller()"
            # replace "xlwings_test_project.xlsm" with the name of the xlsm file
            if "xw.Book(" in line:
                line = line.replace("xlwings_test_project.xlsm", xlsx_file_name_without_extension + ".xlsm")
            # Write content to second file
            secondfile.write(line)

    print("Copied the Python code from " + display_images_py_file + " to " + target_py_file)
    print("All good!")


if __name__ == '__main__':
    # Read the xlsx file which is the first argument.
    xlsx_file = sys.argv[1]
    # For how to add arguments to the script from command line with sys.argv, see 
    # https://www.tutorialspoint.com/python/python_command_line_arguments.htm
    # The argument with index 0 is the name of the script, 
    # then the one with index 1 is the path to the xlsx file, and so on. 
    # See also:
    # https://stackoverflow.com/a/32750302/5193830
    # https://stackoverflow.com/a/17544474/5193830
    start_project(xlsx_file)