import os
import win32com.client

def convert_docx_to_pdf(docx_path):
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(docx_path)
    pdf_path = docx_path.replace(".docx", ".pdf")
    doc.SaveAs(pdf_path, FileFormat=17)  # 17 is the file format for .pdf
    doc.Close()
    word.Quit()
    return pdf_path

def convert_docs_in_folder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        # Skip the "Obsoletos" folder
        if "Obsoletos" in dirs:
            dirs.remove("Obsoletos")
        for file in files:
            if file.endswith(".docx"):
                docx_path = os.path.join(root, file)
                try:
                    print(f"Converting {docx_path} to .pdf")
                    pdf_path = convert_docx_to_pdf(docx_path)
                    print(f"Converted {docx_path} to {pdf_path}")
                except Exception as e:
                    print(f"Error converting {docx_path}: {e}")

folder_path = r'##' # Path to the folder containing the .doc files (## must be replaced with the actual path)
convert_docs_in_folder(folder_path)
