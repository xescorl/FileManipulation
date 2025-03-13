import os
import pandas as pd
from PIL import Image
from PIL.ExifTags import TAGS

# Ruta del directorio con los archivos de imagen
directory = r"##"

# Lista para almacenar los metadatos
metadata_list_image = []

# Recorremos todos los archivos del directorio
for filename in os.listdir(directory):
    if filename.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.bmp')):
        file_path = os.path.join(directory, filename)

        try:
            # Abrimos la imagen y extraemos los metadatos
            image = Image.open(file_path)
            exif_data = image._getexif()
            info_data = image.info

            # Guardamos los metadatos en un diccionario
            metadata = {
                "Nombre Archivo": filename,
                "Formato": image.format,
                "Modo": image.mode,
                "Tamaño": image.size,
                "Metadatos": {}
            }

            if exif_data:
                for tag, value in exif_data.items():
                    tag_name = TAGS.get(tag, tag)
                    metadata["Metadatos"][tag_name] = value

            if info_data:
                for key, value in info_data.items():
                    metadata["Metadatos"][key] = value

            metadata_list_image.append(metadata)
        except Exception as e:
            print(f"Error procesando {filename}: {e}")

# Guardar los metadatos en archivos CSV
df_image = pd.DataFrame(metadata_list_image)
csv_path_image = os.path.join(directory, "metadatos_imagenes.csv")
df_image.to_csv(csv_path_image, index=False, encoding="utf-8")
print(f"Metadatos de imágenes extraídos y guardados en: {csv_path_image}")
