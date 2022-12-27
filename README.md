[![DOI](https://zenodo.org/badge/557367197.svg)](https://zenodo.org/badge/latestdoi/557367197)

# Overview - What is this about?

How to use the free functionality of the [xlwings](https://www.xlwings.org/) library to integrate Python with Excel for visualizing annotated images with their associated bounding boxes for object annotation workflows in your object detection project.

![excel_main_view](https://user-images.githubusercontent.com/14074269/209698357-5f6dd7ac-147c-4516-8908-463a5abc10ca.jpg)
![image_with_box](https://user-images.githubusercontent.com/14074269/209698388-d6110da9-34f1-461a-9a7c-13dd91a06715.jpg)

The Syrphid image was downloaded from [wikipedia](https://en.wikipedia.org/wiki/Hover_fly#/media/File:ComputerHotline_-_Syrphidae_sp._(by)_(3).jpg)

In our AI object detection project, we used [VGG Image Annotator (VIA)](https://www.robots.ox.ac.uk/~vgg/software/via/) to manually annotate insects in images, that is, manually place a bounding box and general taxa information. However, it is difficult to filter and edit metadata fields with VIA, while Excel is more user friendly for such tasks. Therefore, it was necessary to visualize the annotated images directly from Excel.

This repository provides the tools to view images directly from within Excel, together with the associated bounding box of an annotated object.

From Excel, one can click on any row, and a Python script will read the image path together with the coordinates of the bounding box and display the image in a window together with the bounding box.

# Installation - How to make it work?

## Installation of xlwings addin (for Windows)

- The xlwings addin needs [conda](https://www.anaconda.com/products/distribution/start-coding-immediately) to be installed and with it, Python will also be installed. Follow this tutorial from Anaconda's documentation: [Installing on Windows](https://docs.anaconda.com/anaconda/install/windows/). To test if it is already installed, from the Start menu, open the Anaconda Prompt (or Anaconda Powershell Prompt), then type the command `conda list`. A list of installed packages should appear.
- You also need git installed and you can download the executable file from [here](https://git-scm.com/download/win)
- I assume that Python is installed; if not, you can check this [tutorial](https://www.digitalocean.com/community/tutorials/install-python-windows-10);
- To check if Python is installed, in Anaconda Prompt type `python` and you should see something like this:
    ```sh
    C:\Users\your_user_name> python
    
    Python 3.10.9 (tags/v3.10.9:1dd9be6, Dec  6 2022, 20:01:21) [MSC v.1934 64 bit (AMD64)] on win32
    Type "help", "copyright", "credits" or "license" for more information.
    >>> 
    ```
- Install the Excel add-in with the commands:
    ```sh
    pip install xlwings
    xlwings addin install
    # You should see something like:
    # xlwings version: 0.28.5
    # Successfully installed the xlwings add-in!
    ```
- In a/any Excel file, you need to enable the macro options: menu File > Options > Trust Center > Trust Center Settings > Macro Settings > “Enable all macros...". For safety reasons, you can disable this after you are done with your work.
- In a/any Excel file enable the xlwings add-in: menu File > Options > Add-ins > button "Go..." (usually at the bottom, to the right of "Manage: Excel Add-ins"); Click “Browse” and search for a path similar to this one `C:\Users\you_user_name\AppData\Roaming\Microsoft\Excel\XLSTART`; Select the file `xlwings.xlam`; OK; YES (if asked to replace the existing file); OK again;
- At his point, you should see a new menu/tab named "xlwings" in the Excel file (after the Help menu/tab); 

Optionally, if you prefer to run the boxel tool from command line, then clone this repository at your favorite location, for example, to `C:\Users\your_user_name\Documents` and then install the dependencies:
```sh
cd C:\Users\%USERNAME%\Documents
git clone https://github.com/valentinitnelav/img-with-box-from-excel
cd img-with-box-from-excel
pip install -r requirements.txt
# You'll get a series of messages and finally should see something like:
# Successfully installed Pillow-9.3.0 et-xmlfile-1.1.0 numpy-1.23.5 openpyxl-3.0.10 etc.
# If they are already installed, then you will see messages like:
# "Requirement already satisfied: ..."
```
## Excel data structure: 

- In our case, the annotation data can be stored in an Excel file (we'll call it further `data_file.xlsx`) in which each row represents information about a single bounding box.
- The first row of the Excel file must act as the header of the data and must not have empty cells within cells with data (each column should have a name);
- Each row should have at least the following columns (exactly these names) so that the tool works without any other adjustments (see image above):
    - `windows_img_path`: string type, the full/absolute path to the image, e.g. `I:\data\field-images\2021-07-06\Centaurea-scabiosa-01\IMG_0377.JPG`;
    - `id_box`: integer, the id of each box as recorded by the [VGG Image Annotator (VIA)](https://www.robots.ox.ac.uk/~vgg/software/via/); 
    - `x`, `y`, `width` & `height` integer type columns as given by VIA; these are the bounding box coordinates, where `x` & `y` represent the upper left corner (the origin).

## Run the tool

### Via the Graphical User Interface (GUI)

Download the file `boxcel.exe` from this repository and save it anywhere on your computer.
Or if you already cloned this repository (see above), then navigate with Windows Explorer to the folder where you cloned this repository and there you find `boxcel.exe`. This file can be moved anywhere on your computer and it will still execute the Python tool with the associated GUI without the need of installing Python. However, the 3rd party software, xlwings addin, requires conda to be installed (see above).

Open `boxcel.exe`. Click the button "Browse & execute". Choose the desired Excel file (e.g. `data_file.xlsx`). The file must respect the mentioned data structure - see above. Click "Open", Then the tool will create the needed Python file (e.g. `data_file.py`) in the same folder with the Excel file. If successful, you will be notified with a message like "All good! Python code generated. Choose another file or close the application."

Open the Excel file, click on any cell, go to the xlwings menu, and press the green play button named "Run main". The tool will read the current row information with the image path from the column `windows_img_path`, the `id_box` and the box coordinates from `x`, `y`, `width` & `height` columns, and will display the image with its bounding box and a label with the box id.
It will work on any sheet in your `data_file.xlsx` file as long as it can find the required columns mentioned above and they contain valid values.

Alternatively, if all Python dependencies are in place you can also run the tool like this:

Navigate with Windows Explorer to the folder where you cloned this repository, then to the `img-with-box-from-excel\src\boxcel` and right-click on the `gui.py` file and choose "Open with..." then Python. This will start the GUI.

If the above fails, then you can also start the GUI script from command line:
```sh
# In a terminal/command line navigate to the cloned repository and then to the src/boxcel folder
cd C:\Users\%USERNAME%\Documents\img-with-box-from-excel\src\boxcel
python gui.py
```
### Via command line

- Assuming you have a file called `data_file.xlsx` (with the requirements from above), to make it ready to run with this xlwings tool, in the Anaconda Prompt (or Anaconda Powershell Prompt) do this:
```sh
# In a terminal/command line navigate to the cloned repository and then to the src/boxcel folder
cd C:\Users\%USERNAME%\Documents\img-with-box-from-excel\src\boxcel
# Execute the start_project.py which takes as argument the path to your Excel file:
python start_project.py path\to\your\data_file.xlsx # or python3 ...
# Example:
# python start_project.py C:\Users\%USERNAME%\Downloads\data_file.xlsx

# You should see something like:
# xlwings version: 0.28.5
# Copied the Python code from C:\Users\%USERNAME%\Documents\img-with-box-from-excel\src\boxcel\display_images.py to C:\Users\%USERNAME%\Downloads\data_file.py
# All good!
```
This just created the `data_file.py` in the same folder with `data_file.xlsx`.

Additional resources for [xlwings](https://docs.xlwings.org/en/latest/) and the xlwings add-in:

- [How to Supercharge Excel With Python](https://towardsdatascience.com/how-to-supercharge-excel-with-python-726b0f8e22c2) by Costas Andreou;
- The official xlwings website with installation documentation - [here](https://docs.xlwings.org/en/latest/installation.html)

# How to cite this repository?

If this work helped you in any way and would like to cite it, you can do so with a DOI from Zenodo, like:

> Valentin Ștefan. (2022). Use `xlwings` to integrate Python with Excel VBA for visualizing images with their corresponding bounding boxes for your AI project. (v1.0.0). Zenodo. https://doi.org/10.5281/zenodo.7250165
