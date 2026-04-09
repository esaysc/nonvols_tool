import fitz  # PyMuPDF
import easyocr
import pandas as pd
import os

def full_ocr_easyocr(pdf_path, output_excel):
    # Initialize EasyOCR
    # lang=['ch_sim', 'en'] for Simplified Chinese and English
    reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
    
    doc = fitz.open(pdf_path)
    all_rows = []
    
    for page_num in range(len(doc)):
        print(f"Processing page {page_num + 1}/{len(doc)}...")
        page = doc[page_num]
        
        # Convert PDF page to image
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) 
        img_path = f"temp_page_{page_num}.png"
        pix.save(img_path)
        
        # Perform OCR
        result = reader.readtext(img_path)
        
        if result:
            # result format: [([[x,y],...], text, confidence), ...]
            page_data = []
            for item in result:
                box = item[0]
                text = item[1]
                y_top = box[0][1]
                x_left = box[0][0]
                page_data.append({'text': text, 'x': x_left, 'y': y_top})
            
            # Sort by Y
            page_data.sort(key=lambda x: x['y'])
            
            # Row grouping
            if page_data:
                rows = []
                current_row = [page_data[0]]
                for i in range(1, len(page_data)):
                    if abs(page_data[i]['y'] - current_row[-1]['y']) < 15:
                        current_row.append(page_data[i])
                    else:
                        current_row.sort(key=lambda x: x['x'])
                        rows.append(" ".join([item['text'] for item in current_row]))
                        current_row = [page_data[i]]
                current_row.sort(key=lambda x: x['x'])
                rows.append(" ".join([item['text'] for item in current_row]))
                
                for r in rows:
                    all_rows.append({'page': page_num + 1, 'content': r})
        
        # Clean up
        if os.path.exists(img_path):
            os.remove(img_path)

    # Save to Excel
    df = pd.DataFrame(all_rows)
    df.to_excel(output_excel, index=False)
    print(f"Full OCR results (EasyOCR) saved to {output_excel}")

if __name__ == "__main__":
    pdf_file = '贵平2022年大豆发放花名册扫描_1.pdf'
    full_ocr_easyocr(pdf_file, "full_ocr_easyocr_output.xlsx")
