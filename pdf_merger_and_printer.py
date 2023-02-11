import os
from PyPDF2 import PdfReader, PdfMerger


def merge_pdf_files(folder_path):
    """Merge all PDF files in a folder into a single file"""
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    if not pdf_files:
        return

    # output_folder = os.path.join(os.path.split(folder_path)[0], 'results')
    output_folder = os.path.join(os.getcwd(), 'results')
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, folder_path.replace('/', '_').replace('\\', '_').replace('_Users_jemsp_PycharmProjects_NBA_Excel_Attainment_','') + '.pdf')

    merger = PdfMerger()
    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(folder_path, pdf_file)
        with open(pdf_file_path, 'rb') as f:
            merger.append(PdfReader(f))
    merger.write(output_file)


def process_folder(folder_path):
    """Process all sub-folders and files in a folder"""
    merge_pdf_files(folder_path)
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isdir(item_path):
            process_folder(item_path)
            print(f"Procesing : {item_path}")


def main(root_folder):
    """Main function to process all folders and sub-folders starting from root_folder"""
    process_folder(root_folder)


main("C:/Users/jemsp/PycharmProjects/NBA_Excel_Attainment/Result_Folder")
