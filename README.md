# Overview

Integrate Python with Excel VBA for visualizing images with their corresponding bounding boxes for your AI project. 
Why still Excel? Because it is familiar to many people and still very powerful and user-friendly, especially when you need to add many annotation metadata fields and filter data in these fields/columns.

This repository consists of a single Python script that allows image visualization from within Excel, together with the associated bounding box of an annotated object).

From within Excel, one can click on any row and the script will read the image path and the coordinates of the bounding box and display the image in a window together with the box placed on the object of interest.

In our project, we used [VGG Image Annotator (VIA)](https://www.robots.ox.ac.uk/~vgg/software/via/) to manually annotate insects in images: place a bounding box and add various taxa information and additional metadata. However, it is difficult to filter metadata (annotation data) with VIA, so our taxonomists feel at ease using Excel with the annotated dataset.

![xlwings-01](https://user-images.githubusercontent.com/14074269/197849882-fc5bba75-7ac2-48e9-b0be-c67fd173342e.jpg)
![xlwings-02](https://user-images.githubusercontent.com/14074269/197849897-1cb8b94e-bf4b-4aed-a6ae-cd9bb4b23f4d.jpg)
The Syrphid image was downloaded from [wikipedia](https://en.wikipedia.org/wiki/Hover_fly#/media/File:ComputerHotline_-_Syrphidae_sp._(by)_(3).jpg)

# Installation / How to make it work?

One should have an Excel file where each row stores information about a single bounding box.
Each row should have these columns so that the provided Python script works without any other adjustments:

- `windows_img_path`: string, the full/absolute path to the image, e.g. `I:\data\field-img\2021-07-06\Centaurea-scabiosa-01\IMG_0377.JPG`;
- `id_box`: integer, the id of each box as given by VGG-VIA; 
- `x`, `y`, `width` & `height` integer type columns as given by VGG-VIA; these are the bounding box coordinates, where x & y represent the upper left corner (the origin);

The Python script works only if each row in the Excel file corresponds to a single bounding box.
This script should be stored anywhere together with its corresponding xlsm file.
One needs to provide the xlsm file name under `if __name__ == "__main__":` in the script.
Also, rename the script to match the name of the xlsm file as well.

One needs to install [xlwings](https://docs.xlwings.org/en/latest/) and the xlwings add-in. See also these two tutorials for the installation:

- [How to Supercharge Excel With Python](https://towardsdatascience.com/how-to-supercharge-excel-with-python-726b0f8e22c2) by Costas Andreou;
- The official xlwings website with installation documentation - [here](https://docs.xlwings.org/en/latest/installation.html)

**Here are the important steps (for Windows):**

- From keyboard: Windows button + R;
- Type `cmd` then hit Enter; this will start the cmd.exe on Windows (is a terminal where you can write instructions for the computer to execute)
- I assume that Python is installed; if not can check this [tutorial](https://www.digitalocean.com/community/tutorials/install-python-windows-10)
- To check if Python is installed type `where python` and you should see something like `C:\Python38\python.exe`;
- Install xlwings with the command `pip install xlwings`; If all goes well, you should see something like:
```sh
xlwings version: 0.28.3
Successfully installed the xlwings add-in!
```
- To install all needed dependencies you can try the command `pip install "xlwings[all]"` as suggested official xlwings documentation page, [here](https://docs.xlwings.org/en/latest/installation.html#optional-dependencies);
- Install the Excel add-in with `xlwings addin install`
- In any Excel file, you need to enable the macro options: menu File > Options > Trust Center > Trust Center Settings > Macro Settings > “Enable all macros..."
- Ccreate a template project with the command `xlwings quickstart project_name` (in the terminal, use `cd` to set the needed path, for example, `cd Documents`). This creates the folder `project_name` which contains two files (you can rename them, but should have the same name): 
  - project_name.xlsm
  - project_name.py
- In the project_name.xlsm, enable the xlwings add-in by pressing the keys ALT+L+H; click “Browse” and search for this path `C:\Users\you_user_name\AppData\Roaming\Microsoft\Excel\XLSTART`; select the file xlwings.xlam; OK; YES (if asked to replace the existing file); OK again;
- At his point, you should see a new menu/tab named "xlwings" in any Excel file (after the Help menu/tab); 
- Copy your Excel data (see the minimum column requirements above) into the project_name.xlsm file;
- Copy the content or download the Python script from this repository (img-with-box-from-excel.py) and replace project_name.py. Rename if needed so that it matches the name of the xlsm file;
- Provide the xlsm file name under `if __name__ == "__main__":` in the Python script
```python
if __name__ == "__main__":
    xw.Book("project_name.xlsm").set_mock_caller() # !!! Add your file name
    main()
```
- All set. Click in the Excel file on any cell, go to menu xlwings and press the green play button named “Run main”. The script will read the current row information with the image path from the column `windows_img_path`, the `id_box` and the box coordinates from `x`, `y`, `width` & `height` columns, and will display the image with its bounding box and a label with the box id.



