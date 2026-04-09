import fitz  # PyMuPDF
from paddleocr import PaddleOCR
import pandas as pd
import os

def full_ocr_process(pdf_path, output_excel):
    # Initialize PaddleOCR
    # Use default settings, but ensure CPU mode if needed
    # lang='ch' for Chinese
    ocr = PaddleOCR(lang='ch')
    
    doc = fitz.open(pdf_path)
    all_rows = []
    
    for page_num in range(len(doc)):
        print(f"Processing page {page_num + 1}/{len(doc)}...")
        page = doc[page_num]
        
        # Convert PDF page to image with high resolution
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3)) 
        img_path = f"temp_page_{page_num}.png"
        pix.save(img_path)
        
        # Perform OCR
        result = ocr.ocr(img_path)
        
        if result and result[0]:
            # Sort results by Y coordinate to group by rows
            # line format: [ [[x1,y1], [x2,y2], [x3,y3], [x4,y4]], (text, confidence) ]
            lines = result[0]
            # Sort by top-y
            lines.sort(key=lambda x: x[0][0][1])
            
            page_data = []
            for line in lines:
                box = line[0]
                text = line[1][0]
                y_top = box[0][1]
                x_left = box[0][0]
                page_data.append({'text': text, 'x': x_left, 'y': y_top})
            
            # Simple row grouping logic
            if page_data:
                rows = []
                current_row = [page_data[0]]
                for i in range(1, len(page_data)):
                    # If Y difference is small, consider same row
                    if abs(page_data[i]['y'] - current_row[-1]['y']) < 20:
                        current_row.append(page_data[i])
                    else:
                        # Sort current row by X
                        current_row.sort(key=lambda x: x['x'])
                        rows.append(" ".join([item['text'] for item in current_row]))
                        current_row = [page_data[i]]
                # Add last row
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
    print(f"Full OCR results saved to {output_excel}")

if __name__ == "__main__":
    # Disable oneDNN to avoid the previous error
    os.environ['FLAGS_enable_mkldnn'] = '0'
    os.environ['FLAGS_use_mkldnn'] = '0'
    pdf_file = '贵平2022年大豆发放花名册扫描_1.pdf'
    full_ocr_process(pdf_file, "full_ocr_output.xlsx")
