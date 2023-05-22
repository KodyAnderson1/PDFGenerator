# YAML to Word Doc Converter

This is a simple script to convert a YAML file to a Word document. It is written in Python 3.6 and uses
the [python-docx](https://python-docx.readthedocs.io/en/latest/) library.

The document contains a title, a heading with the student's name, course, and date, and a list of sections as headings.
Each section starts on a new page. The document is saved with a file name that includes the lowercase last name of the
student and the lab number. For example, `doe_lab01.docx`.

## Installation

1. Install Python 3.x from [python.org](https://www.python.org/downloads/).
2. Install the required libraries by running `pip install -r requirements.txt` in the project directory.
3. **Optional:** Create a virtual environment by running `python -m venv venv` in the project directory before
   installing the libraries.

## Usage

1. Create a YAML file with the following structure:

```yaml
title:
  labNumber: Lab 02
  labName: Chapter 02 Lab
heading:
  Name: Kody Anderson
  Course: CTS 4348 Linux System Administration
sections:
  - On workstation, run the lab start cli-review script
  - Display the current time in 24-hour clock time
  - What kind of file is /home/student/zcat? Is it readable by humans?
  - Use the wc command and Bash shortcuts to display the size of zcat
  - Display the first 10 lines of zcat.
  - Display the last 10 lines of the zcat file
  - Repeat the previous command exactly with three or fewer keystrokes
  - Repeat the previous command, but use the -n 20 option...
  - Use the shell history to run the date +%r command again.
  - On workstation, run the lab grade cli-review...
  - Execute the ls -l command to show the file created in the previous step.
  - On workstation, run the lab finish cli-review script to complete the lab

docxDirectory: './'
imgDirectory: 'C:\Users\Owner\Pictures\Screenshots\LinuxAdmin\Lab02\'

```

2. Save the YAML file in the project directory as `input.yaml`.
3. Run `python main.py` in the project directory.
4. The Word document will be saved in the project directory as `[lastname]_[lab_##].docx`.





