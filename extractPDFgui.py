import tkinter as tk
from tkinter import filedialog
import pdfplumber
import pandas as pd
import os

# Function to extract tables from PDF
def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract tables from the page
            page_tables = page.extract_tables()
            if page_tables:
                tables.extend(page_tables)
    return tables

# Function to clean and organize tables
def clean_and_organize_tables(tables):
    cleaned_tables = []
    for table in tables:
        # Remove empty rows and columns
        table = [row for row in table if any(cell and cell.strip() for cell in row)]
        if table:
            # Use the first row as headers
            headers = table[0]
            data = table[1:]
            df = pd.DataFrame(data, columns=headers)
            cleaned_tables.append(df)
    return cleaned_tables

# Function to extract tables and save to Excel
def extract_tables_to_excel(pdf_path, excel_path):
    # Extract tables from PDF
    tables = extract_tables_from_pdf(pdf_path)
    
    # Clean and organize tables
    cleaned_tables = clean_and_organize_tables(tables)
    
    # Save each table to a separate sheet in the Excel file
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for i, df in enumerate(cleaned_tables):
            sheet_name = f"Table_{i+1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    
    print(f"Tables have been extracted and saved to {excel_path}")

# Function to browse and select PDF file
def browse_pdf():
    filename = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    pdf_path_entry.delete(0, tk.END)
    pdf_path_entry.insert(0, filename)
    
    # Set the default Excel path based on the PDF filename
    default_excel_path = os.path.splitext(filename)[0] + '.xlsx'
    
    # Check if the Excel file already exists and rename if necessary
    if os.path.exists(default_excel_path):
        base, extension = os.path.splitext(default_excel_path)
        counter = 1
        new_excel_path = f"{base}_{counter}{extension}"
        while os.path.exists(new_excel_path):
            counter += 1
            new_excel_path = f"{base}_{counter}{extension}"
        default_excel_path = new_excel_path
    
    excel_path_entry.delete(0, tk.END)
    excel_path_entry.insert(0, default_excel_path)

# Function to browse and select Excel file
def browse_excel():
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    excel_path_entry.delete(0, tk.END)
    excel_path_entry.insert(0, filename)

# Function to run the extraction process
def run_extraction():
    pdf_path = pdf_path_entry.get()
    excel_path = excel_path_entry.get()
    extract_tables_to_excel(pdf_path, excel_path)

# Create the main window
root = tk.Tk()
root.title("PDF to Excel Extractor")

# Create and place the PDF path widgets
tk.Label(root, text="PDF Path:").grid(row=0, column=0, padx=10, pady=10)
pdf_path_entry = tk.Entry(root, width=50)
pdf_path_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_pdf).grid(row=0, column=2, padx=10, pady=10)

# Create and place the Excel path widgets
tk.Label(root, text="Excel Path:").grid(row=1, column=0, padx=10, pady=10)
excel_path_entry = tk.Entry(root, width=50)
excel_path_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=browse_excel).grid(row=1, column=2, padx=10, pady=10)

# Create and place the Extract button
tk.Button(root, text="Extract", command=run_extraction).grid(row=2, columnspan=3, pady=20)

# Run the application
root.mainloop()
