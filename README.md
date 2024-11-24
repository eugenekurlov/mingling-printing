<!-- omit in toc -->
# Mingling Printing

A tkinter-based GUI application that allows users to combine multiple image or PDF files into a single document, customize output settings, and print seamlessly with tailored configurations. Designed to streamline file management and printing tasks for enhanced efficiency.

---
<!-- omit in toc -->
## Table of Contents

- [Overview](#overview)
- [Features](#features)
  - [General Features](#general-features)
  - [Image Settings](#image-settings)
  - [PDF Settings](#pdf-settings)
  - [Printer Settings](#printer-settings)
- [System Requirements](#system-requirements)
- [Get started](#get-started)
- [Usage](#usage)
  - [Initial Window](#initial-window)
  - [Files Configuration Window from images](#files-configuration-window-from-images)
  - [Files Configuration Window from pdf files](#files-configuration-window-from-pdf-files)
  - [Printer Settings Window](#printer-settings-window)
- [Building Executable](#building-executable)
- [Disclaimer](#disclaimer)
- [License](#license)

---

## Overview

If you use Windows 10 operating system, you can select several images or PDF files in one folder, right-click mouse button on the selection, select "print". But there some limitations. In images case changing of printing settings can only be done through "options", "printer properties" in the "Print images" window that opens, and you cannot change the order of files:
it depends on which file you right-clicked first from the selected ones. It is not allowed to set printer setting while print PDF files this way.

This application enables users to combine and print multiple images or PDF files as a single document with customizable settings. The interface is straightforward and user-friendly and allows for easy file selection, configuration, and printing.

---

## Features

### General Features

- Select files from any directory through the file manager.
- Combine images or PDF files into a single document.
- Save combined files on the disk for reuse or send them directly to a printer.
- Configurable printing settings, including duplex mode and number of copies.

### Image Settings

- Rotate images manually or enable "Best Orientation" for automatic layout.
- Choose grid layout (rows and columns) for multiple images on one page.
- Save resulting PDF with options to define orientation (portrait, landscape, or auto).

### PDF Settings

- Specify pages for every single PDF file to include.
- Merge pages from multiple PDFs into one document.
- Save output to desktop or IO buffer.

### Printer Settings

- Default printer selection with the ability to choose others.
- Set the number of copies.
- Enable duplex printing with configurable "long edge" or "short edge" modes.

---

## System Requirements

- **Operating System**: Windows 8 or later (because Python 3.12 cannot be used on Windows 7 or earlier).
- **PDF Reader**: A PDF viewer (e.g., Adobe Acrobat Reader, Foxit Reader, Sumatra PDF and so on) must be installed to open the generated PDF files.
- **Python Version**: Python 3.12 or later.

---

## Get started

1. Clone this repository:

   ```bash
   git clone https://github.com/eugenekurlov/mingling-printing.git
   ```

2. Navigate to the project directory:

   ```bash
   cd mingling-printing
   ```

3. Create a virtual environment.

   On Windows:

   ```cmd
   python -m venv venv
   ```

   or

   ```cmd
   py -m venv venv
   ```

4. Activate the virtual environment:

   For Command Prompt:

   ```cmd
   venv\Scripts\activate
   ```

   For PowerShell:

   ```PowerShell
   .\venv\Scripts\Activate.ps1
   ```

5. Install dependencies into virtual environment:

   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### Initial Window

1. Upon launching, the user can choose to:
   - **Images**: Opens the [**Files Configuration Window**](#files-configuration-window-from-images) from images.
   - **PDFs**: Opens the [**Files Configuration Window**](#files-configuration-window-from-pdf-files) for `.pdf` files.

---

### Files Configuration Window from images

1. Use the **Select Images** button to choose image files using file manager.
2. Supported formats: `.jpg`, `.jpeg`, `.png`.
3. The thumbnail version of each image and its title are shown after selection.
4. Selected file can be removed from the list by pressing `x` sign button.
5. User can change files' order by pressing "up" and "down" arrows.
6. Configure options:
   - **Specify angles**. Checkbox button, not selected by default, meaning no image rotation.
   - - If selected, the entry fileds appear. It accept positive or negative ineger value defining angle to rotate every single image. Otherwise it means zero rotating angle.
   - - If selected, the **Best Orientation** option becomes disabled.
   - <a id="output_path_id"></a> **Specify Output Path**. Checkbox button, not selected by default. If active, the file manager open to save the resulting file to the disk. If it's none, then the the resulting file will be has `io.BytesIO()` type.
   - **Best Orientation**. Checkbox button, not selected by default, meaning no image rotation to best fit the sheet's grid borders based on grid's cell width and height, and image's width and height.
   - - If selected, the **Specify angles** option becomes disabled and `"Auto"` orientation becomes enabled.
   - - If selected and if **Orientation** is `"Auto"` and **Multiple Pages (Grid)** options equal 1 for row and column, then the whole pdf page fit its orientation to image's width and height.
   - **Multiple Pages (Grid)**. Checkbox button, not selected by default, meaning 1 images per sheet. Allows to put the several images on the sheet (by rows and columns).
   - **Orientation**. Defines `"Portrait"`, `"Landscape"`, or `"Auto"` (with grid disabled) pdf pages orientations.
   - - if **Best Orientation** is selected and if **Orientation** is `"Auto"` and **Multiple Pages (Grid)** options are not 1, then the parameter is set to `"Portrait"`
   - **Page margin**. Entry filed with default value is 0. It is a distance in millimeters from each page edge to images usable area. Must be integer. Autochecking the `int` value was implemented in tkinter gui. If it is empty, then value is set to 0.
   - **Image margin**. Entry filed with default value is 0. It is a distance in millimeters between image in grid cell and its borders. Must be integer. Autochecking the `int` value was implemented in tkinter gui. If it is empty, then value is set to 0.
7. Generate the PDF using the **Generate PDF** button.
8. <a id="print_button_id"></a>After generating the PDF, the **Print** button appears for printing options. Open [**Printer Settings**](#printer-settings-window) window.

---

### Files Configuration Window from pdf files

1. Use the **Select PDF Files** button to choose PDF files using file manager.
2. Supported formats: `.pdf`.
3. Select specific pages from the files using entry field to include in the final document. If the field is empty, all pages from a file will be added. Input format is a string without whitespaces. The string should contain onlt digits, comma and hyphen signs. The string cannot be started with hyphen; hyphen and comma sign cannot be next to each other. To select the range of pages the start and finish range value should be separated with hyphen. An example of input format: `1-3,7,10-12`. The hyphen can be placed the last after page number: this means the range from the page to the last page of the file.
4. [Save](#output_path_id) or [print](#print_button_id) the resulting merged PDF using similar steps (№6 and 8) as in the **Image Settings Window**.

---

### Printer Settings Window

1. Accessed after clicking the **Print** button from either Images or PDFs window.
2. Configure:
   - **Printer selection**. Combobox. Default value is the OS default printer.
   - **Number of copies**. Spinbox. The number of how much times to print file.
   - - **Collate**. Checkbox button, is hidden by default. It appears only when **Number of copies** is bigger then 1.
   - **Duplex option**. Checkbox button, not selected by default.
   - - If **Duplex option** is selected, the **flip side** frame appears with `"Long Edge"` (by default) or `"Short Edge"` options for printer device to flip each pages during printing.

---

## Building Executable

To create an executable file for the application:

1. Ensure all dependencies are installed (`requirements.txt`).
2. Install `pyinstaller` in your Python environment:

   ```bash
   pip install pyinstaller==6.11.1
   ```

3. In the command prompt, navigate to the project directory and run:

   ```bash
   pyinstaller --onefile --windowed --icon=printer.ico --add-data "printer.ico;." app.py
   ```

4. The `.exe` file will be generated in the `dist/` folder.

---

## Disclaimer

Windows is a registered trademark of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft Corporation.

---

## License

Mingling printing

Copyright (&copy;) 2024 eugenekurlov

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.