# This script extracts the student ID and any information in a PASS plan relating to extra time in exams and in-class tests. 
# PASS plans are in PDF format, and older LSPs are in .docx format. This script extracts data from both types of file.
# Results are saved in a .xlsx file.

import os
import re
import fitz
import pandas as pd
import sys
import openpyxl
import tkinter as tk
from tkinter import filedialog
import docx
from docx import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.oxml import parse_xml
from docx.shared import Pt

root = tk.Tk()
root.withdraw()
doc_folder = filedialog.askdirectory()

if not doc_folder:
    print("No folder selected.")
    exit()

output_excel_file = "/Users/mattpsychology/Documents/PASS Plan Extract.xlsx"

# matches all text found in the PDF between "Exams and In Class Tests" and "Advisor", "Assessments", or "Advice".
pattern_tests = re.compile(r"Exams and In Class Tests([\s\S]*?)(?:Advisor|Assessments|Advice)", re.IGNORECASE)

# matches all text (the student ID) found in the PDF between "Student Id" and "College".
pattern_student_id = re.compile(r"Student Id([\s\S]*?)College", re.IGNORECASE)

# matches all text found in the .docx file between "In Examinations" and "Copies of this".
pattern_lsp_tests = re.compile(r"In Examinations([\s\S]*?)(?:Copies of this)", re.IGNORECASE)

# matches all text (the student ID) found in the .docx file where there are eight consecutive digits.
pattern_lsp_student_id = re.compile(r"\d{8}", re.IGNORECASE)

def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    text = ""
    for page in doc:
        text += page.get_text("text") + "\n"
    return text

def extract_text_from_docx(docx_path):
    doc = docx.Document(docx_path)
    text = []

    for element in doc.element.body:
        if isinstance(element, CT_P):
            text.append(element.text)
        elif isinstance(element, CT_Tbl):
            for row in element.findall(".//w:tr", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                row_text = []
                for cell in row.findall(".//w:tc", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    cell_text = "".join([t.text or "" for t in cell.findall(".//w:t", namespaces={'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'})])
                    row_text.append(cell_text.strip())
                text.append(" | ".join(row_text))
    return "\n".join(text)

results = []

for filename in os.listdir(doc_folder):
    file_path = os.path.join(doc_folder, filename)
    text = ""

    if filename.lower().endswith(".pdf"):
        text = extract_text_from_pdf(file_path)
        match_tests = pattern_tests.findall(text)
        match_student_id = pattern_student_id.findall(text)
    elif filename.lower().endswith(".docx"):
        text = extract_text_from_docx(file_path)
        match_tests = pattern_lsp_tests.findall(text)
        match_student_id = pattern_lsp_student_id.findall(text)
    else:
        continue

    extracted_tests = match_tests[0].strip() if match_tests else "No match found"
    extracted_student_id = match_student_id[0].strip() if match_student_id else "No match found"

    results.append({
        "Student ID": extracted_student_id,
        "Extracted Text": extracted_tests
    })

df = pd.DataFrame(results)
df.to_excel(output_excel_file, index=False, engine="openpyxl")

print(f"Search completed. Results saved to {output_excel_file}")
