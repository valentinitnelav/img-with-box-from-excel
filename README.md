[![DOI](https://zenodo.org/badge/557367197.svg)](https://zenodo.org/badge/latestdoi/557367197)

# Overview - What is this about?

How to use the free functionality of the [xlwings](https://www.xlwings.org/) library to integrate Python with Excel for visualizing annotated images with their associated bounding boxes for object annotation workflows in your object detection project.

![xlwings-01](https://user-images.githubusercontent.com/14074269/197849882-fc5bba75-7ac2-48e9-b0be-c67fd173342e.jpg)
![xlwings-02](https://user-images.githubusercontent.com/14074269/197849897-1cb8b94e-bf4b-4aed-a6ae-cd9bb4b23f4d.jpg)

The Syrphid image was downloaded from [wikipedia](https://en.wikipedia.org/wiki/Hover_fly#/media/File:ComputerHotline_-_Syrphidae_sp._(by)_(3).jpg)

Why use Excel for image annotation workflows?
This is because it is familiar to many people, especially when you need to add many annotation metadata fields and filter data.

In our AI object detection project, we used [VGG Image Annotator (VIA)](https://www.robots.ox.ac.uk/~vgg/software/via/) to manually annotate insects in images, that is, manually place a bounding box and add record taxa information together with custom metadata. However, it is difficult to filter and edit metadata fields with VIA, while Excel is more user friendly for such tasks. Therefore, it was necessary to visualize the annotated images directly from Excel.

This repository provides the tools to view images directly from within Excel, together with the associated bounding box of an annotated object.

From within Excel, one can click on any row, and a Python script will read the image path together with the coordinates of the bounding box and display the image in a window together with the box placed on the object of interest.

# Installation - How to make it work?

Excel data structure: 

- One should have the annotation data stored in an Excel file (`*.xlsm` not `*.xlsx`; see details below) in which each row represents information about a single bounding box. The tool works only if each row in the Excel file corresponds to a single bounding box.
- The first row of the Excel file must act as the header of the data and must not have empty cells within cells with data (each column should have a user defined name);
- Each row should have the following columns so that the tool works without any other adjustments:
    - `windows_img_path`: string type, the full/absolute path to the image, e.g. `I:\data\field-images\2021-07-06\Centaurea-scabiosa-01\IMG_0377.JPG`;
    - `id_box`: integer, the id of each box as recorded by the [VGG Image Annotator (VIA)](https://www.robots.ox.ac.uk/~vgg/software/via/); 
    - `x`, `y`, `width` & `height` integer type columns as given by VIA; these are the bounding box coordinates, where `x` & `y` represent the upper left corner (the origin).


## Here are the important steps (for Windows):

### Installation of dependencies (do only once)

- Start Command Prompt (`cmd.exe`) on Windows (is a terminal where you can write instructions for the computer to execute). From keyboard: Windows button + R then type `cmd`, then hit Enter. Or type directly `cmd` in the search box of the Start menu in Windows OS, then hit Enter.
- I assume that Python is installed; if not, you can check this [tutorial](https://www.digitalocean.com/community/tutorials/install-python-windows-10);
- To check if Python is installed, type `where python` and you should see something like `C:\Python38\python.exe`;
- Clone this repository at your favorite location, for example, to `C:\Users\your_user_name\Documents` and then install the dependencies:
```sh
cd C:\Users\your_user_name\Documents
git clone https://github.com/valentinitnelav/img-with-box-from-excel
cd img-with-box-from-excel
pip install -r requirements.txt
```
- Install the Excel add-in with the command 
```sh
xlwings addin install
```
- In a/any Excel file, you need to enable the macro options: menu File > Options > Trust Center > Trust Center Settings > Macro Settings > “Enable all macros..."

### Run the tool

- Assuming you have a file called `data_file.xlsx`, you would use it like this on the xlsx file:
```sh
# In a terminal/command line navigate to the cloned repository and then to the src/boxcel folder
cd C:\Users\your_user_name\Documents\img-with-box-from-excel\src\boxcel
# Execute the start_project.py which takes as argument the path to your Excel file:
python start_project.py path\to\your\data_file.xlsx # or python3 ...
# Example:
# python start_project.py C:\Users\vs66tavy\Downloads\hymenoptera_sample.xlsx
```
This creates a folder within the folder where `data_file.xlsx` is located with 2 files named after the Excel file: `data_file.xlsm` and `data_file.py`. All data from `data_file.xlsx` were copied to `data_file.xlsm`.

- Open `data_file.xlsm` and enable the xlwings add-in: menu File > Options > Add-ins > button "Go..." (usually at the bottom, to the right of "Manage: Excel Add-ins"); Click “Browse” and search for a path similar to this one `C:\Users\you_user_name\AppData\Roaming\Microsoft\Excel\XLSTART`; Select the file `xlwings.xlam`; OK; YES (if asked to replace the existing file); OK again;
- At his point, you should see a new menu/tab named "xlwings" in the Excel file (after the Help menu/tab); 

All set. Click in the Excel file on any cell, go to the xlwings menu, and press the green play button named "Run main". The tool will read the current row information with the image path from the column `windows_img_path`, the `id_box` and the box coordinates from `x`, `y`, `width` & `height` columns, and will display the image with its bounding box and a label with the box id.
It will work on any sheet in your `data_file.xlsm` file as long as it can find the minimum required columns mentioned above and they contain valid values.


Additional resources for [xlwings](https://docs.xlwings.org/en/latest/) and the xlwings add-in:

- [How to Supercharge Excel With Python](https://towardsdatascience.com/how-to-supercharge-excel-with-python-726b0f8e22c2) by Costas Andreou;
- The official xlwings website with installation documentation - [here](https://docs.xlwings.org/en/latest/installation.html)

# How to cite this repository?

If this work helped you in any way and would like to cite it, you can do so with a DOI from Zenodo, like:

> Valentin Ștefan. (2022). Use `xlwings` to integrate Python with Excel VBA for visualizing images with their corresponding bounding boxes for your AI project. (v1.0.0). Zenodo. https://doi.org/10.5281/zenodo.7250165
