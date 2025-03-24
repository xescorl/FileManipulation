import os
import re

def rename_edition_files(folder_path):
    for root, dirs, files in os.walk(folder_path):
        # Skip the "Obsoletos" folder
        if "Obsoletos" in root:
            continue
        
        for filename in files:
            # Use regex to find and replace [number]ªEd with Ed[number]
            new_filename = re.sub(r'(\d+)ªEd', r'Ed\1', filename)
            
            if filename != new_filename:
                print(f'Renaming {filename} to {new_filename}')
                os.rename(os.path.join(root, filename), os.path.join(root, new_filename))

folder_path = r'C:\Users\ftur.DOMINIOEXPO\Dropbox\2_ISO EXPOCOM\0_POLITICA, ORGANIGRAMA'
rename_edition_files(folder_path)