import os
import glob
import tempfile
import time
from PyPDF2 import PdfMerger, PdfReader
from pdf2docx import Converter
import win32com.client
import logging
from tqdm import tqdm
import re
from io import StringIO
from datetime import datetime

# Suppress INFO messages
logging.getLogger('win32com').setLevel(logging.WARNING)
logging.getLogger('pdf2docx').setLevel(logging.WARNING)

# Create a StringIO object to capture all output
output_capture = StringIO()


def log_output(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_message = f"[{timestamp}] {message}"
    print(formatted_message)
    print(formatted_message, file=output_capture)


def display_intro():
    border = "+" + "-" * 78 + "+"
    title = "Document Merger, Converter, and Word Counter"
    creator = "Script created by Christopher D. van der Kaay, Ph.D. (August 15, 2024)"

    log_output(border)
    log_output(f"|{title:^78}|")
    log_output(f"|{creator:^78}|")
    log_output(border)
    log_output("|" + " " * 78 + "|")
    log_output("| This script performs the following tasks:                                     |")
    log_output("| 1. Converts all DOC and DOCX files in a specified input folder to PDF format. |")
    log_output("| 2. Merges all PDF files (including the newly converted ones) into a single    |")
    log_output("|    PDF.                                                                       |")
    log_output("| 3. Converts the merged PDF back into a DOCX file.                             |")
    log_output("| 4. Counts the frequency of user-specified words in the final document.        |")
    log_output("|    (Case-insensitive matching)                                                |")
    log_output("| 5. Generates a timestamped audit log file of all operations.                  |")
    log_output("|" + " " * 78 + "|")
    log_output("| The script will prompt you for:                                               |")
    log_output("| 1. The input folder containing the files to be processed                      |")
    log_output("| 2. The name of the final output DOCX file  e.g., merged_doc.docx              |")
    log_output("| 3. Words to search for in the final document                                  |")
    log_output("|" + " " * 78 + "|")
    log_output("| The final DOCX file and audit log will be saved in your Documents folder.     |")
    log_output("| The audit log (audit_log.txt) contains a detailed record of all operations    |")
    log_output("| performed by the script, including timestamps and any errors encountered.     |")
    log_output(border)
    input("Press Enter to begin...")


def convert_docx_to_pdf(docx_file, output_dir):
    word = win32com.client.Dispatch('Word.Application')
    doc = word.Documents.Open(docx_file)
    output_pdf = os.path.join(output_dir, os.path.splitext(os.path.basename(docx_file))[0] + '.pdf')
    doc.SaveAs(output_pdf, FileFormat=17)  # FileFormat=17 is for PDF
    doc.Close()
    word.Quit()
    return output_pdf


def merge_pdfs(pdf_files, output_pdf):
    merger = PdfMerger()
    for pdf in tqdm(pdf_files, desc="Merging PDFs", unit="file"):
        merger.append(pdf)
    merger.write(output_pdf)
    merger.close()


def convert_pdf_to_docx(pdf_file, output_file):
    cv = Converter(pdf_file)
    cv.convert(output_file, start=0, end=None)
    cv.close()


def ensure_pdf_format(files, output_dir):
    pdf_files = []
    for file in tqdm(files, desc="Converting to PDF", unit="file"):
        try:
            if file.lower().endswith(('.docx', '.doc')):
                pdf_file = convert_docx_to_pdf(file, output_dir)
                pdf_files.append(pdf_file)
                log_output(f"Converted {file} to PDF")
            elif file.lower().endswith('.pdf'):
                pdf_files.append(file)
                log_output(f"Added existing PDF: {file}")
            else:
                log_output(f"Skipping unsupported file format: {file}")
        except Exception as e:
            log_output(f"Error processing file {file}: {e}")
    return pdf_files


def get_words_to_search():
    log_output("Enter words to search for (one per line). Press Enter on a blank line to finish:")
    words = []
    while True:
        word = input().strip().lower()
        if not word:
            break
        words.append(word)
        log_output(f"Added search word: {word}")
    return words


def count_word_frequency(pdf_file, words_to_search):
    word_counts = {word: 0 for word in words_to_search}
    reader = PdfReader(pdf_file)
    for page in tqdm(reader.pages, desc="Counting words", unit="page"):
        text = page.extract_text().lower()
        for word in words_to_search:
            word_counts[word] += len(re.findall(r'\b' + re.escape(word) + r'\b', text, re.IGNORECASE))
    return word_counts


def main(input_dir, output_file):
    try:
        log_output(f"Analysis started for input directory: {input_dir}")
        with tempfile.TemporaryDirectory() as temp_dir:
            input_files = glob.glob(os.path.join(input_dir, "*.docx")) + \
                          glob.glob(os.path.join(input_dir, "*.doc")) + \
                          glob.glob(os.path.join(input_dir, "*.pdf"))
            if not input_files:
                log_output("No DOC, DOCX, or PDF files found in the directory.")
                return

            log_output("Ensuring all files are in PDF format...")
            pdf_files = ensure_pdf_format(input_files, temp_dir)

            merged_pdf = os.path.join(temp_dir, "merged.pdf")
            log_output("Merging PDF files...")
            merge_pdfs(pdf_files, merged_pdf)

            log_output("Converting merged PDF to DOCX...")
            convert_pdf_to_docx(merged_pdf, output_file)

            words_to_search = get_words_to_search()
            if words_to_search:
                word_counts = count_word_frequency(merged_pdf, words_to_search)
                log_output("\nWord Frequency Results:")
                for word, count in word_counts.items():
                    log_output(f"'{word}': {count} occurrences")

        log_output(f"Merged document saved as {output_file}")
    except Exception as e:
        log_output(f"An error occurred: {e}")
    finally:
        time.sleep(2)


if __name__ == "__main__":
    start_time = datetime.now()
    log_output(f"Script execution started at {start_time}")

    display_intro()
    input_folder = input("Enter the path to the folder containing the files: ").strip()
    if not os.path.isdir(input_folder):
        log_output("The provided path is not a valid directory.")
    else:
        output_docx = input("Enter the name for the final output DOCX file (e.g., final_document.docx): ").strip()
        if not output_docx.endswith('.docx'):
            output_docx += ".docx"
        output_docx = os.path.join(os.path.expanduser("~"), "Documents", output_docx)
        main(input_folder, output_docx)
        log_output(f"File should be saved at: {os.path.abspath(output_docx)}")
        if os.path.exists(output_docx):
            log_output("File successfully created!")
        else:
            log_output("ERROR! File was not created. Please check the permissions and try again.")

        # Generate audit log file
        audit_log_file = os.path.join(os.path.dirname(output_docx), "audit_log.txt")
        with open(audit_log_file, "w") as log_file:
            log_file.write(output_capture.getvalue())
        log_output(f"Audit log file created at: {audit_log_file}")

    end_time = datetime.now()
    log_output(f"Script execution completed at {end_time}")
    log_output(f"Total execution time: {end_time - start_time}")

    input("\nPress Enter to close the window...")
