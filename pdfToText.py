import re
import os
import PyPDF2
import yaml


def parse_pdf_file(file_path):
    """Parse the PDF file and return the text."""
    pdf_file_obj = open(file_path, 'rb')
    pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
    text = ""
    for page_num in range(len(pdf_reader.pages)):
        page_obj = pdf_reader.pages[page_num]
        text += page_obj.extract_text()
    pdf_file_obj.close()
    return text


def extract_after_string(text, specified_string):
    """Extract the text after the specified string."""
    try:
        extracted_text = text.split(specified_string, 1)[1]
    except IndexError:
        extracted_text = ""
    return extracted_text


def filter_lines_starting_with_number(lines):
    """Filter lines that start with a number and remove the starting number."""
    filtered_lines = []
    for line in lines:
        if re.match(r'^\d', line):
            filtered_line = re.sub(r'^\d+\s*', '', line)
            filtered_line = re.sub(r'^\.', '-', filtered_line)
            filtered_lines.append(filtered_line)
    return filtered_lines


def write_to_file(file_path, lines):
    """Write lines to a file."""
    with open(file_path, 'w') as f:
        for line in lines:
            f.write("%s\n" % line)


def extract_lines_starting_with_number_after_string(pdf_file_path, specified_string, output_file_path):
    """Main function to parse the PDF file for a specified string and then extract every line after that starting with a number."""
    # Parse the PDF file
    text = parse_pdf_file(pdf_file_path)

    # Extract text after the specified string
    extracted_text = extract_after_string(text, specified_string)

    # Filter lines that start with a number and remove the starting number
    lines = extracted_text.split('\n')
    filtered_lines = filter_lines_starting_with_number(lines)

    # Write lines to a file
    write_to_file(output_file_path, filtered_lines)


if __name__ == "__main__":
    with open("input.yaml", 'r') as stream:
        try:
            config = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)

    pdf_file_path = config['pdf_file_path']
    specified_string = config['specified_string']
    output_file_path = config['output_file_path']
    extract_lines_starting_with_number_after_string(pdf_file_path, specified_string, output_file_path)
