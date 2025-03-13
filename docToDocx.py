import os
import win32com.client

def convert_doc_to_docx(doc_path):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(doc_path)
    docx_path = doc_path + "x"
    doc.SaveAs(docx_path, FileFormat=16)  # 16 is the file format for .docx
    doc.Close()
    word.Quit()
    return docx_path

def convert_docs_in_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        # Skip the "Obsoletos" folder
        if "Obsoletos" in dirs:
            dirs.remove("Obsoletos")
        for file in files:
            if file.endswith(".doc") and not file.endswith(".docx"):
                doc_path = os.path.join(root, file)
                try:
                    print(f"Converting {doc_path} to .docx")
                    docx_path = convert_doc_to_docx(doc_path)
                    print(f"Deleting original .doc file: {doc_path}")
                    os.remove(doc_path)
                except Exception as e:
                    print(f"Error converting {doc_path}: {e}")

folder_path = r'##' # Path to the folder containing the .doc files (## must be replaced with the actual path)
convert_docs_in_folder(folder_path)