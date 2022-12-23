# This script helps a Excel user to visualize images which have their full/absolute paths
# stored in a column (here named 'windows_img_path'). Each images has also 4 coordinates for 
# the bounding box (x, y, width, height) as given by the VGG Image Annotator (VIA),
# see https://www.robots.ox.ac.uk/~vgg/software/via/
# The script works only if each row in the Excel file corresponds to a single bounding box.

# This script should be stored anywhere together with its corresponding xlsm file.
# Need to provide the file name under `if __name__ == "__main__":` below.
# Rename the script to match the name of the xlsm file as well.

import xlwings as xw
import pandas as pd
import PIL
from PIL import Image, ImageDraw, ImageFont


def display_img():
    """
    This function does the followings:
    - reads info from the calling/current workbook, sheet, cell;
    it reads from the following user-defined columns (must be named exactly like this):
    windows_img_path, id_box, x, y, width & height (the last 4 are the bounding box coordinates, 
    where x & y are the origin = upper left corner)
    - displays the image with the bounding box in temporary window
    """
    
    # Get the calling/current workbook
    wb = xw.Book.caller()
    # Get the calling/current sheet
    sht = wb.sheets.active
    # Return the row id of the calling/current selected cell
    row_id = wb.app.selection.row

    # Read each current (selected) row as a data frame.
    # First read the header of the sheet (1st row).
    # It only works properly as long as there are no empty cells within cells
    # with values in the first row / header of the sheet.
    cols = sht.range(1,1).expand(mode='right').value
    # Then read each current row based on the length of the header (cols is a list here)
    line = sht.range(cell1=(row_id,1), cell2=(row_id,len(cols))).value
    # Create the 1-row data frame which contains the info stored in the selected row
    df = pd.DataFrame([line], columns=cols)

    # Read the box coordinates and prepare them for PIL.ImageDraw.Draw.rectangle()
    x0 = df['x'][0]
    y0 = df['y'][0]
    x1 = df['x'][0] + df['width'][0]
    y1 = df['y'][0] + df['height'][0]
    coord = [x0, y0, x1, y1]

    # Read also the box id; this can be useful to have on the image.
    id_box = str(int(df['id_box'][0]))
    font = ImageFont.truetype('arial', 40)
    # font can create problems because it could be that the path to font file 
    # needs to be specified.

    # Read image path from Excel
    img_path = df['windows_img_path'][0]
    # Open image, draw the box & the box id
    with Image.open(img_path) as im:
        draw = ImageDraw.Draw(im)
        # https://pillow.readthedocs.io/en/stable/reference/ImageDraw.html#PIL.ImageDraw.ImageDraw.rectangle
        draw.rectangle(xy=coord, outline='Red', width=3)
        # Place the box id a bit more towards the upper left corner
        draw.text((x0-10, y0-10), id_box, fill='Blue', font=font, stroke_fill='White', stroke_width=3)
        im.show(title=img_path)  
        # title doesn't work with `show`, 
        # see https://github.com/python-pillow/Pillow/issues/5739
        # This is because PIL opens a temporary PNG file for the current image file.
        # Also, "Users of the library should use a context manager or call Image.Image.close() 
        # on any image opened with a filename or Path object to ensure that the underlying file is closed."
        # https://pillow.readthedocs.io/en/stable/reference/open_files.html#proposed-file-handling
        im.close() # I am unsure about this, but it doesn't break the code :)


# When you hit the "Run main" (play green button) from the xlwings tab, 
# it will run the main function declared below, which calls display_img()
# Or you can create a button to which you can assign the VBA Subroutine.
# For example, the SampleCall() subroutine from the template module (Module 1)
#  that was created with the xlwings_test_project.xlsm by the cmd line `xlwings quickstart <name>`
def main():
    display_img()

# Change below "xlwings_test_project.xlsm" with your xlsm file name.
if __name__ == "__main__":
    xw.Book("xlwings_test_project.xlsm").set_mock_caller() # !!! Add your xlsm file name !!!
    main()
