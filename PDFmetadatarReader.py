import os
import pandas as pd
from PyPDF2 import PdfReader

# Ruta del directorio con los archivos PDF
directory = r"##"

# Lista para almacenar los metadatos
metadata_list = []

# Recorremos todos los archivos del directorio y sus subdirectorios
for root, _, files in os.walk(directory):
    for filename in files:
        if filename.lower().endswith(".pdf"):
            file_path = os.path.join(root, filename)

            try:
                # Abrimos el PDF y extraemos los metadatos
                with open(file_path, "rb") as file:
                    pdf = PdfReader(file)
                    info = pdf.metadata

                    # Guardamos los metadatos en un diccionario
                    metadata = {
                        "Nombre Archivo": filename,
                        "Título": info.get("/Title", "No disponible"),
                        "Autor": info.get("/Author", "No disponible"),
                        "Asunto": info.get("/Subject", "No disponible"),
                        "Palabras Clave": info.get("/Keywords", "No disponible"),
                        "Creador": info.get("/Creator", "No disponible"),
                        "Productor": info.get("/Producer", "No disponible"),
                        "Fecha de Creación": info.get("/CreationDate", "No disponible"),
                        "Fecha de Modificación": info.get("/ModDate", "No disponible"),
                        "Número de Páginas": len(pdf.pages),
                    }

                    metadata_list.append(metadata)
            except Exception as e:
                print(f"Error procesando {filename}: {e}")

# Guardar los metadatos en un archivo CSV
df = pd.DataFrame(metadata_list)
csv_path = os.path.join(directory, "metadatos_pdf.csv")
df.to_csv(csv_path, index=False, encoding="utf-8")

print(f"Metadatos extraídos y guardados en: {csv_path}")
