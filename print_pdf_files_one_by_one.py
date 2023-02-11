import os
import win32api
import win32print
import time
from PyPDF2 import PdfReader


def print_pdf_file(pdf_file_path):
    """Print a PDF file to the default printer in duplex mode"""
    with open(pdf_file_path, 'rb') as f:
        win32api.ShellExecute(0, "print", pdf_file_path, None, ".", 0)
        win32print.SetPrinter(win32print.GetDefaultPrinter(), 2)


def print_pdf_files(folder_path):
    """Print all PDF files in a folder, one by one, to the default printer in duplex mode"""
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    if not pdf_files:
        return

    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(folder_path, pdf_file)
        print_pdf_file(pdf_file_path)


def main(root_folder):
    """Main function to print all PDF files in all folders and sub-folders starting from root_folder to the default printer in duplex mode"""
    for item in os.listdir(root_folder):
        item_path = os.path.join(root_folder, item)
        print(item_path)
        if os.path.isdir(item_path):
            print_pdf_files(item_path)

        time.sleep(4)


main("C:/Users/jemsp/PycharmProjects/NBA_Excel_Attainment/results")
