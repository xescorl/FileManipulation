import os
import pandas as pd
from docx import Document

# Ruta del directorio con los archivos Word
directory = r"##"

# Lista para almacenar los metadatos
metadata_list = []

# Recorremos todos los archivos del directorio y sus subdirectorios
for root, _, files in os.walk(directory):
    for filename in files:
        if filename.lower().endswith(".docx"):
            file_path = os.path.join(root, filename)

            try:
                # Abrimos el archivo Word y extraemos los metadatos
                doc = Document(file_path)
                core_properties = doc.core_properties

                # Guardamos los metadatos en un diccionario
                metadata = {
                    "Nombre Archivo": filename,
                    "Título": core_properties.title if core_properties.title else "No disponible",
                    "Autor": core_properties.author if core_properties.author else "No disponible",
                    "Asunto": core_properties.subject if core_properties.subject else "No disponible",
                    "Palabras Clave": core_properties.keywords if core_properties.keywords else "No disponible",
                    "Comentarios": core_properties.comments if core_properties.comments else "No disponible",
                    "Categoría": core_properties.category if core_properties.category else "No disponible",
                    "Fecha de Creación": core_properties.created if core_properties.created else "No disponible",
                    "Fecha de Modificación": core_properties.modified if core_properties.modified else "No disponible",
                    "Número de Páginas": len(doc.element.xpath('//w:sectPr'))  # Approximation based on section breaks
                }

                metadata_list.append(metadata)
            except Exception as e:
                print(f"Error procesando {filename}: {e}")

# Guardar los metadatos en un archivo CSV
df = pd.DataFrame(metadata_list)
csv_path = os.path.join(directory, "metadatos_word.csv")
df.to_csv(csv_path, index=False, encoding="utf-8")

print(f"Metadatos extraídos y guardados en: {csv_path}")
