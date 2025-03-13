import os
from pathlib import Path
import PyPDF2

def remove_metadata(file_path):
    try:
        with open(file_path, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            writer = PyPDF2.PdfWriter()
            for page in reader.pages:
                writer.add_page(page)
            writer.add_metadata({})

            with open(file_path, 'wb') as output_file:
                writer.write(output_file)

        print(f"Metadata removed from {file_path}")
    except Exception as e:
        print(f"Failed to remove metadata from {file_path}: {e}")

def clean_metadata_in_directory(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = Path(root) / file
            if file_path.suffix.lower() == '.pdf':
                remove_metadata(file_path)

if __name__ == "__main__":
    directory = r'##'
    clean_metadata_in_directory(directory)
