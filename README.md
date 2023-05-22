# Lab Report Generation Script Usage Guide

This is a Python script that generates a lab report in .docx format with an option to convert it to .pdf. The script
uses a YAML file for input data and can embed images from a given directory into the report.

## Getting Started

1. Install the required Python packages: `pip install -r requirements.txt`
2. Save your script and the YAML file in the same directory.
3. Prepare your images:
    - Store the images that you want to include in your report in a **single** directory.
    - Ensure all the images are in **PNG** format.
    - The images will be added to the report in alphanumeric order, so name them appropriately.

4. Update the YAML file:

- Provide the lab number and name.
- Input your name and course.
- List out the sections of the lab.
- Specify the directories for the input images and output .docx and .pdf files.

## Running the Script

1. Navigate to the directory where your script and YAML file are stored.
2. Run the script with Python: `python main.py`
3. Once the Word (.docx) document is generated, the script will prompt you to convert the document to PDF. Press 'Enter'
   to continue or 'CTRL+C' to exit.

**Note:** If you want to add a table of contents, you need to add it manually to the Word document. Make sure to close
the Word document before you press 'Enter' to convert it to a PDF.

## YAML File Format

Your YAML file should follow this format:

```yaml
title:
labNumber: <Your Lab Number> # MUST BE A STRING, THEN A SPACE, THEN A NUMBER. I.E. Lab 02
labName: <Your Lab Name>

heading:
  Name: <Your Name> # MUST BE A STRING, THEN A SPACE, THEN A STRING. I.E. John Doe
  Course: <Your Course>
sections:
  - <List of>
  - <Section Titles>
  - <In the Order>
  - <You want them>
  - <To Appear>
docxDirectory: <Directory where the .docx file will be saved>
pdfDirectory: <Directory where the .pdf file will be saved>
imgDirectory: <Directory where your .png images are stored>
```

```text
Docs will be formatted as: 

< labNumber: labName >
-----------------------------
< Your Name >
< Your Course >
< Today's Date >
```

## Additional Notes

- This script works best on Windows and macOS systems. On Linux and other operating systems, the script may not
  terminate Word processes correctly.
- The script only supports .png images. If you have images in another format, convert them to .png before running the
  script.
- By default, the images will be resized to 6.5 inches wide in the Word document to fit within the default margins (1
  inch on each side). If you want to change this, adjust the IMG_WIDTH_INCHES constant at the top of the script.
- The script sorts and adds images based on their alphanumeric order. Name your images accordingly, e.g., '
  01_first_image.png', '02_second_image.png', etc.

