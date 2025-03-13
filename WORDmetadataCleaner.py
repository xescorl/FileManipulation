import os
from pathlib import Path
from docx import Document
from datetime import datetime

def remove_metadata(file_path):
    try:
        doc = Document(file_path)
        core_properties = doc.core_properties

        # Clear metadata
        core_properties.author = ""
        core_properties.subject = ""
        core_properties.keywords = ""
        core_properties.comments = ""
        core_properties.category = ""
        core_properties.created = datetime.min
        core_properties.modified = datetime.min

        doc.save(file_path)
        print(f"Metadata removed from {file_path}")
    except Exception as e:
        print(f"Failed to remove metadata from {file_path}: {e}")

def clean_metadata_in_directory(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = Path(root) / file
            if file_path.suffix.lower() == '.docx':
                remove_metadata(file_path)

if __name__ == "__main__":
    directory = r'##'
    clean_metadata_in_directory(directory)
