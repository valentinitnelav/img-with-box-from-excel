# Run the script like this:
# python3 start_project.py path/to/file.xlsx

import sys
import os

def start_project(xlsx_file):
    """
    This function writes the needed code to the *.py file corresponding to the *.xlsx file.
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
    path_to_dir_boxcel = sys.path[0] # this should return path/to/img-with-box-from-excel/src/boxcel
    display_images_py_file = os.path.join(path_to_dir_boxcel, "display_images.py")

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
    # Read the xlsx file which is the frist argument.
    xlsx_file = sys.argv[1]
    # For how to add arguments to the script from command line with sys.argv, see 
    # https://www.tutorialspoint.com/python/python_command_line_arguments.htm
    # The argument with index 0 is the name of the script, 
    # then the one with index 1 is the path to the xlsx file, and so on. 
    # See also:
    # https://stackoverflow.com/a/32750302/5193830
    # https://stackoverflow.com/a/17544474/5193830
    start_project(xlsx_file)