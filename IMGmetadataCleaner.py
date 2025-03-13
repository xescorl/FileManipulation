import os
from pathlib import Path
from PIL import Image

def remove_metadata(file_path):
    try:
        image = Image.open(file_path)
        data = list(image.getdata())
        image_without_metadata = Image.new(image.mode, image.size)
        image_without_metadata.putdata(data)
        image_without_metadata.save(file_path)
        print(f"Metadata removed from {file_path}")
    except Exception as e:
        print(f"Failed to remove metadata from {file_path}: {e}")

def clean_metadata_in_directory(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            file_path = Path(root) / file
            if file_path.suffix.lower() in ['.jpg', '.jpeg', '.png', '.tiff', '.bmp']:
                remove_metadata(file_path)

if __name__ == "__main__":
    directory = r'##'
    clean_metadata_in_directory(directory)
