import os
import re
from tkinter import Tk, Button, Label, filedialog
from docx import Document
import fitz  

def clean_text(text):
    cleaned_text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', text)
    return cleaned_text


def convert_pdf_to_docx(pdf_path, docx_path):
    doc = fitz.open(pdf_path)
    document = Document()
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text()
        cleaned_text = clean_text(text)
        document.add_paragraph(cleaned_text)
    document.save(docx_path)

def select_input_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    input_files_label.config(text=", ".join(file_paths))
    return file_paths

def convert_to_docx():
    file_paths = input_files_label.cget("text").split(", ")
    for file_path in file_paths:
        if file_path.endswith('.pdf'):
            output_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
            if output_path:
                convert_pdf_to_docx(file_path, output_path)
                print(f'{os.path.basename(file_path)} converted to {os.path.basename(output_path)} successfully.')
    status_label.config(text="Conversion complete.")
root = Tk()
root.title("PDF to DOCX Converter")
root.update_idletasks()
root.geometry(f"{root.winfo_width()}x{root.winfo_height()}")
input_files_label = Label(root, text="Select input files")
input_files_label.grid(row=0, column=0)
status_label = Label(root, text="")
status_label.grid(row=2, column=0)
input_files_button = Button(root, text="Browse", command=select_input_files)
input_files_button.grid(row=0, column=1)
convert_button = Button(root, text="Convert", command=convert_to_docx)
convert_button.grid(row=1, column=0, columnspan=2)
root.mainloop()