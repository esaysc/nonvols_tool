import fitz  # PyMuPDF
from paddleocr import PaddleOCR
import pandas as pd
import os

def ocr_pdf_to_excel(pdf_path, output_excel):
    # Initialize PaddleOCR with minimal arguments
    # Try to disable mkldnn via constructor if possible, or just use defaults
    ocr = PaddleOCR(lang='ch', use_gpu=False)
    
    doc = fitz.open(pdf_path)
    all_results = []
    
    for page_num in range(len(doc)):
        print(f"Processing page {page_num + 1}/{len(doc)}...")
        page = doc[page_num]
        
        # Convert PDF page to image
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2)) # Increase resolution
        img_path = f"temp_page_{page_num}.png"
        pix.save(img_path)
        
        # Perform OCR
        result = ocr.ocr(img_path)
        
        if result and result[0]:
            for line in result[0]:
                box = line[0]
                text = line[1][0]
                confidence = line[1][1]
                
                y_center = (box[0][1] + box[2][1]) / 2
                x_center = (box[0][0] + box[1][0]) / 2
                all_results.append({
                    'page': page_num,
                    'text': text,
                    'x': x_center,
                    'y': y_center,
                    'conf': confidence
                })
        
        if os.path.exists(img_path):
            os.remove(img_path)

    df_raw = pd.DataFrame(all_results)
    df_raw.to_excel("ocr_raw_debug.xlsx", index=False)
    print("Raw OCR results saved to ocr_raw_debug.xlsx")
    
    return df_raw

if __name__ == "__main__":
    # Force CPU and disable mkldnn via environment variables
    os.environ['FLAGS_enable_mkldnn'] = '0'
    os.environ['FLAGS_cpu_deterministic'] = '1'
    os.environ['PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK'] = 'True'
    
    pdf_file = '贵平2022年大豆发放花名册扫描_1.pdf'
    try:
        ocr_pdf_to_excel(pdf_file, "ocr_output.xlsx")
    except Exception as e:
        print(f"OCR failed: {e}")
        print("Attempting to extract text directly as fallback...")
        # Fallback to simple text extraction if OCR fails completely
        doc = fitz.open(pdf_file)
        text_content = []
        for page in doc:
            text_content.append(page.get_text())
        with open("pdf_text_fallback.txt", "w", encoding="utf-8") as f:
            f.write("\n".join(text_content))
        print("Fallback text saved to pdf_text_fallback.txt")
