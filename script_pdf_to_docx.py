import fitz  # PyMuPDF
from docx import Document

def pdf_to_word_fitz(pdf_file, word_file):
    # Open the PDF file
    pdf_document = fitz.open(pdf_file)
    doc = Document()

    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        
        doc.add_paragraph(text)

    # Save the document
    doc.save(word_file)
    print(f"Conversion complete! {word_file} has been created.")

if __name__ == "__main__":
    pdf_file = 'Unit 1_Ch_1.pdf'
    word_file = 'output1.docx'
    
    pdf_to_word_fitz(pdf_file, word_file)
