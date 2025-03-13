import os

def change_extension(folder_path, old_ext, new_ext):
    for filename in os.listdir(folder_path):
        if filename.endswith(old_ext):
            base = os.path.splitext(filename)[0]
            new_filename = base + new_ext
            print(f'Changing {filename} to {new_filename}')
            os.rename(os.path.join(folder_path, filename), os.path.join(folder_path, new_filename))
        elif not os.path.splitext(filename)[1]:  # Check if there is no extension
            new_filename = filename + new_ext
            print(f'Adding extension to {filename} to make it {new_filename}')
            os.rename(os.path.join(folder_path, filename), os.path.join(folder_path, new_filename))

folder_path = r'##'
change_extension(folder_path, '.txt', '.py')