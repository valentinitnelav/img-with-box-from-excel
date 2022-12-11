# Run the script like this:
# python3 start_project.py path/to/file.xlsx

import sys
import os
import pandas as pd
from openpyxl import load_workbook


# Read the xlsx file as the frist argument.
xlsx_file = sys.argv[1]
# For how to add arguments to the script from command line with sys.argv, see 
# https://www.tutorialspoint.com/python/python_command_line_arguments.htm
# The argument with index 0 is the name of the script, 
# then the one with index 1 is the path to the xlsx file, and so on. 
# See also:
# https://stackoverflow.com/a/32750302/5193830
# https://stackoverflow.com/a/17544474/5193830

# Test if the xlsx file exists
if not os.path.isfile(xlsx_file):
    print("The file " + xlsx_file + " does not exist.")
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


# Read all the sheets from the xlsx file.
try:
    xl = pd.ExcelFile(xlsx_file, engine='openpyxl')
except:
    print("Error trying to read from " + xlsx_file)
    sys.exit()


# Create a template project with the command xlwings quickstart project_name
# The command is run from the directory where the xlsx file is located.
# It creates a folder with the name of the xlsx file without the extension.
os.chdir(path_to_xlsx_file)
os.system("xlwings quickstart " + xlsx_file_name_without_extension)


# Write the data from xl to the xlwings template xlsm file for each sheet.
# The xlsm file is located in the folder with the name of the xlsx file.
# https://stackoverflow.com/a/42375263/5193830
xlsm_file = os.path.join(path_to_xlsx_file, 
                         xlsx_file_name_without_extension,
                         xlsx_file_name_without_extension + ".xlsm")
book = load_workbook(xlsm_file)
# Delete the first empty sheet from the xlsm template file.
book.remove(book["Sheet1"])

writer = pd.ExcelWriter(xlsm_file, engine='openpyxl')
writer.book = book
for sheet_name in xl.sheet_names:
    print("Writing sheet " + sheet_name + " to xlsm file " + xlsm_file)
    df = xl.parse(sheet_name)
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Move the sheet "_xlwings.conf" to the end of the xlsm file.
# This sheet was created together with Sheet1 (deleted above) when `xlwings quickstart` was run.
sheet = book["_xlwings.conf"]
book.remove(sheet)
book._sheets.append(sheet)

writer.close()

# Copy the Python code for the created xlwings template xlsm file.
# This consists in adding the name of the xlsm file to the script display_images.py
# Read the path to the folder where this script file is located
current_file = os.path.realpath(__file__)
path_to_current_file = os.path.dirname(current_file)
display_images_py_file = os.path.join(path_to_current_file, "display_images.py")
# Check also this:
# display_images_py_file = pkgutil.get_data(__package__, "src/display_images.py")
# from https://stackoverflow.com/a/51724506/5193830

target_py_file = os.path.join(path_to_xlsx_file,
                              xlsx_file_name_without_extension, 
                              xlsx_file_name_without_extension + ".py")

with open(display_images_py_file,'r') as firstfile, open(target_py_file,'w') as secondfile:
    # Read content from first file
    for line in firstfile:  
        # For line "xw.Book("xlwings_test_project.xlsm").set_mock_caller()"
        # replace "xlwings_test_project.xlsm" with the name of the xlsm file
        if "xw.Book(" in line:
            line = line.replace("xlwings_test_project.xlsm", xlsx_file_name_without_extension + ".xlsm")
        # Write content to second file
        secondfile.write(line)

print("Copied the Python code from " + display_images_py_file + " to " + target_py_file)
print("All good!")