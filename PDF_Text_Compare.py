import os
import tkinter as tk
from tkinter import filedialog
from typing import List

import PyPDF2
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

def extract_text_from_pdf(pdf_file_path: str) -> str:
    """Extracts text from a PDF file."""
    with open(pdf_file_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text()
    return text


def calculate_similarity(text1: str, text2: str) -> float:
    """Calculates the cosine similarity between two text strings."""
    vectorizer = TfidfVectorizer()
    tfidf_matrix = vectorizer.fit_transform([text1, text2])
    similarity_score = cosine_similarity(tfidf_matrix[0:1], tfidf_matrix[1:2])[0][0]
    return similarity_score


def compare_pdfs_with_source(source_file: str, input_folder: str, output_file: str) -> None:
    """Compares PDFs in a folder to a source PDF and writes the results to an output file."""
    source_text = extract_text_from_pdf(source_file)

    with open(output_file, 'w') as outfile:
        outfile.write("File,Similarity Score\n")
        for filename in os.listdir(input_folder):
            if filename.endswith(".pdf"):
                full_path = os.path.join(input_folder, filename)
                try:
                    input_text = extract_text_from_pdf(full_path)
                    similarity = calculate_similarity(source_text, input_text)
                    outfile.write(f"{filename},{similarity}\n")
                    print(f"Compared {filename}, Similarity: {similarity:.2f}")
                except Exception as e:
                    print(f"Error processing {filename}: {e}")


def compare_pdfs_in_folder(input_folder: str, output_file: str) -> None:
    """Compares all PDFs within a folder and writes the results to an output file."""
    pdf_files = [f for f in os.listdir(input_folder) if f.endswith(".pdf")]
    
    with open(output_file, 'w') as outfile:
        outfile.write("File 1,File 2,Similarity Score\n")
        for i in range(len(pdf_files)):
            for j in range(i+1, len(pdf_files)):
                file1_path = os.path.join(input_folder, pdf_files[i])
                file2_path = os.path.join(input_folder, pdf_files[j])
                try:
                    text1 = extract_text_from_pdf(file1_path)
                    text2 = extract_text_from_pdf(file2_path)
                    similarity = calculate_similarity(text1, text2)
                    outfile.write(f"{pdf_files[i]},{pdf_files[j]},{similarity}\n")
                    print(f"Compared {pdf_files[i]} and {pdf_files[j]}, Similarity: {similarity:.2f}")
                except Exception as e:
                    print(f"Error comparing {pdf_files[i]} and {pdf_files[j]}: {e}")


def browse_file(entry: tk.Entry) -> None:
    """Opens a file dialog and sets the selected file path to an Entry widget."""
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                          filetypes=(("PDF files", "*.pdf*"), ("all files", "*.*")))
    entry.delete(0, tk.END)
    entry.insert(0, filename)

def browse_folder(entry: tk.Entry) -> None:
    """Opens a folder dialog and sets the selected folder path to an Entry widget."""
    folder_path = filedialog.askdirectory(initialdir="/", title="Select a Folder")
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

def run_comparison() -> None:
    """Gets input values from the GUI and initiates the PDF comparison."""
    source_file = source_entry.get()
    input_folder = input_entry.get()
    output_file = output_entry.get()

    if source_file:
        compare_pdfs_with_source(source_file, input_folder, output_file)
    else:
        compare_pdfs_in_folder(input_folder, output_file)

# GUI setup
root = tk.Tk()
root.title("PDF Text Comparison")

# Source File
source_label = tk.Label(root, text="Source PDF (Optional):")
source_label.grid(row=0, column=0, padx=10, pady=10)
source_entry = tk.Entry(root, width=50)
source_entry.grid(row=0, column=1, padx=10, pady=10)
source_button = tk.Button(root, text="Browse", command=lambda: browse_file(source_entry))
source_button.grid(row=0, column=2, padx=10, pady=10)

# Input Folder
input_label = tk.Label(root, text="Input Folder:")
input_label.grid(row=1, column=0, padx=10, pady=10)
input_entry = tk.Entry(root, width=50)
input_entry.grid(row=1, column=1, padx=10, pady=10)
input_button = tk.Button(root, text="Browse", command=lambda: browse_folder(input_entry))
input_button.grid(row=1, column=2, padx=10, pady=10)

# Output File
output_label = tk.Label(root, text="Output File:")
output_label.grid(row=2, column=0, padx=10, pady=10)
output_entry = tk.Entry(root, width=50)
output_entry.grid(row=2, column=1, padx=10, pady=10)

# Run Button
run_button = tk.Button(root, text="Run Comparison", command=run_comparison)
run_button.grid(row=3, column=1, pady=20)

root.mainloop()

