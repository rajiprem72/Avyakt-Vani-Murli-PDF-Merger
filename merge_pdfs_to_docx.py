import os
from PyPDF2 import PdfReader
from docx import Document

def merge_pdfs_to_docx(pdf_folder, output_docx):
    # Create a new Word document
    doc = Document()
    
    # Get all PDF files sorted by name (assuming date-wise sequential naming)
    pdf_files = sorted([f for f in os.listdir(pdf_folder) if f.lower().endswith('.pdf')])
    
    if not pdf_files:
        print("No PDF files found in the folder!")
        return

    for pdf_file in pdf_files:
        pdf_path = os.path.join(pdf_folder, pdf_file)
        print(f"Processing: {pdf_file}")
        
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
        except Exception as e:
            print(f"Error processing {pdf_file}: {e}")

    doc.save(output_docx)
    print(f"Successfully created: {output_docx}")

if __name__ == "__main__":
    input_folder = input("Enter the folder path containing PDFs: ").strip()
    output_file = "Merged_Murli_Vani.docx"
    merge_pdfs_to_docx(input_folder, output_file)
