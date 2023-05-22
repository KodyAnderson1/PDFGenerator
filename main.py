import os
import yaml
import datetime
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import argparse
import platform
import subprocess

IMG_WIDTH_INCHES = 6.5  # default image width in inches (can be changed). Default margins are 1 inch on each side


def read_yaml_file(file_path):
    with open(file_path) as yaml_file:
        data = yaml.safe_load(yaml_file)
    return data


def get_png_images(directory):
    return sorted([os.path.join(directory, f) for f in os.listdir(directory) if f.lower().endswith('.png')])


def create_word_document(data, images):
    document = Document()

    # Add title and heading to the first page
    title = document.add_heading(f"{data['title']['labNumber']}: {data['title']['labName']}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    heading = document.add_paragraph(data['heading']['Name'])
    heading.add_run(f"\n{data['heading']['Course']}")
    today = datetime.datetime.today().strftime('%B %d, %Y')
    heading.add_run(f"\n{today}")

    # Add each section to its own page
    for i, section in enumerate(data['sections']):
        document.add_page_break()
        section_title = document.add_heading(section, level=1)
        section_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        section_title.style.font.size = Pt(20)  # Add this line to set font size to 20
        section_page = section_title._element.xpath('w:pPr/w:pageBreakBefore')
        if section_page:
            section_page[0].set('w:val', 'true')

        if images and i < len(images):
            paragraph = document.add_paragraph()
            run = paragraph.add_run()
            run.add_picture(images[i], width=Inches(IMG_WIDTH_INCHES))

    return document


def save_word_document(document, output_file_path):
    document.save(output_file_path)


def get_file_name(data, output_directory):
    name = data['heading']['Name'].split(' ')[1].lower() \
        if len(data['heading']['Name'].split(' ')) > 1 \
        else data['heading']['Name'].split(' ')[0].lower()

    lab_number = data['title']['labNumber'].replace(" ", "").lower()
    return os.path.join(output_directory, f"{name}_{lab_number}")


def convert_to_pdf(input_file_path, pdf_directory):
    # Ensure the file paths are in the correct format
    input_file_path = os.path.abspath(input_file_path)
    output_file_path = os.path.join(pdf_directory, os.path.basename(input_file_path).replace('.docx', '.pdf'))
    print(f"Converting {input_file_path} to {output_file_path}")

    # Kill any existing Word processes
    kill_word_processes()

    # Convert the Word document to PDF
    try:
        import docx2pdf
        docx2pdf.convert(input_file_path, output_file_path)
        print(f"Conversion complete. PDF saved at {output_file_path}")
    except ImportError:
        print("docx2pdf library is not installed. Please install it using 'pip install docx2pdf'.")

    # Kill any remaining Word processes
    kill_word_processes()


def kill_word_processes():
    if platform.system() == 'Windows':
        import psutil
        for proc in psutil.process_iter():
            try:
                if "WINWORD.EXE" == proc.name():
                    proc.kill()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ProcessLookupError):
                pass
    elif platform.system() == 'Darwin':
        subprocess.call(['pkill', 'soffice'])
    else:
        print("Unsupported operating system. Process termination is not supported.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser()

    data = read_yaml_file('input.yaml')

    images = get_png_images(data['imgDirectory'])

    docx_directory = data.get('docxDirectory', '.')
    pdf_directory = data.get('pdfDirectory', '.')

    document = create_word_document(data, images)
    docx_file_path = get_file_name(data, docx_directory) + '.docx'
    save_word_document(document, docx_file_path)

    # Convert the Word document to a PDF
    input(
        'Press enter to convert the Word document to a PDF or CTRL + C to exit.\n'
        'If you want to add a table of contents, you need to add it manually.\n'
        'MAKE SURE YOU CLOSE THE WORD DOCUMENT BEFORE YOU PRESS ENTER\n')
    convert_to_pdf(docx_file_path, pdf_directory)
