import os
import unicodedata

def normalize_and_uppercase(folder_path):
    for root, dirs, files in os.walk(folder_path):
        # Skip the "Obsoletos" folder
        if "Obsoletos" in root:
            continue
        
        for filename in files:
            # Normalize the filename to remove accents
            normalized_name = unicodedata.normalize('NFKD', filename).encode('ASCII', 'ignore').decode('ASCII')
            
            # Split the filename into parts to preserve "Ed" case
            parts = normalized_name.split("Ed")
            # Convert only the parts before and after "Ed" to uppercase
            new_filename = "Ed".join(part.upper() if i != 1 else part for i, part in enumerate(parts))
            
            if filename != new_filename:
                print(f'Renaming {filename} to {new_filename}')
                os.rename(os.path.join(root, filename), os.path.join(root, new_filename))

folder_path = r'C:\Users\ftur.DOMINIOEXPO\Dropbox\2_ISO EXPOCOM\0_POLITICA, ORGANIGRAMA'
normalize_and_uppercase(folder_path)