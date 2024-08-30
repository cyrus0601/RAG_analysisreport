"""
Description: This script detects rotated pages in a PDF document and rotates them.
"""
# %%
import fitz
import pdfplumber

def find_rotated_pages(pdf_path):

    rotated_pages = set()
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages): 
            words = page.extract_words()
            for word in words:
                if not word['upright']:  
                    rotated_pages.add(page_num)  
            
    return sorted(rotated_pages)

def save_pdf_with_rotated_page(pdf_file, pages_to_rotate, output_file):

    pdf_document = fitz.open(pdf_file)
    new_pdf = fitz.open()

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)

        if page_num in pages_to_rotate:
            page.set_rotation(90)  

        new_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)

    new_pdf.save(output_file)
    pdf_document.close()
    new_pdf.close()

def process_pdf(pdf_file, output_file):

    rotated_pages = find_rotated_pages(pdf_file)
    print(f"Rotated pages found: {rotated_pages}")

    save_pdf_with_rotated_page(pdf_file, rotated_pages, output_file)


pdf_file = "" 
output_file = ""  

process_pdf(pdf_file, output_file)
