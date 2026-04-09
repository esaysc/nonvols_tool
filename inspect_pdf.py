import fitz  # PyMuPDF
import pandas as pd

def inspect_pdf(file_path):
    doc = fitz.open(file_path)
    print(f"PDF has {len(doc)} pages.")
    
    for page_num in range(min(3, len(doc))):
        page = doc[page_num]
        text = page.get_text()
        print(f"--- Page {page_num} ---")
        print(text[:1000]) # Print first 1000 chars
        
        # Try to find tables using get_text("words") or similar if needed
        # But first let's see the raw text layout

if __name__ == "__main__":
    inspect_pdf('贵平2022年大豆发放花名册扫描_1.pdf')
