import os
import fitz  # PyMuPDF
from PIL import Image
import pytesseract as tess
import tempfile
import re
import json
from docx import Document
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

TEMP_IMAGE_DIR = tempfile.gettempdir()
INPUT_DIR = "input_docs"
OUTPUT_DIR = "output_docs"
os.makedirs(OUTPUT_DIR, exist_ok=True)

class DocumentProcessor:

    @staticmethod
    def convert_pdf_page_to_image(pdf_path, page_num, dpi=200):
        try:
            doc = fitz.open(pdf_path)
            page = doc.load_page(page_num)
            zoom = dpi / 72
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)

            image_filename = f"{os.path.splitext(os.path.basename(pdf_path))[0]}_page_{page_num+1}.png"
            image_path = os.path.join(TEMP_IMAGE_DIR, image_filename)
            pix.save(image_path)
            doc.close()
            return image_path
        except Exception as e:
            logging.error(f"Error converting page {page_num} in {pdf_path}: {e}")
            return None

    def extract_field(self, text, field_name):
        pattern = rf"{re.escape(field_name)}[:\s]*([^\n\r]*)"
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            return match.group(1).strip()
        return None
    
    def extract_field_multiline(self, text, field_name):
        pattern = rf"{re.escape(field_name)}.*?:\s*(.*?)(?=\n[A-Z][A-Z \(\):]*\n|\Z)"
        match = re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
        if match:
            value = match.group(1).strip()
            value = " ".join(value.split())
            return value
        return None


    def extract_text_from_pdf_ocr(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            all_text = ""
            for page_num in range(doc.page_count):
                image_path = self.convert_pdf_page_to_image(pdf_path, page_num)
                if image_path:
                    img = Image.open(image_path)
                    text = tess.image_to_string(img)
                    all_text += f'\n\n*** Page {page_num + 1} ***\n' + text
                    os.remove(image_path)
            doc.close()
            return all_text
        except Exception as e:
            logging.error(f"Error during OCR extraction from {pdf_path}: {e}")
            return ""

    def process_application(self, app_folder_name, file_paths):
        logging.info(f"Processing application '{app_folder_name}' with {len(file_paths)} files.")
        combined_text = ""
        for file_path in file_paths:
            if file_path.lower().endswith('.pdf'):
                combined_text += self.extract_text_from_pdf_ocr(file_path)
            else:
                logging.warning(f"Unsupported file format {file_path}, skipping.")

        # Field to be extracted * Edit here accordingly
        extracted = {
            "applicant_name": self.extract_field(combined_text, "NAME OF INSTITUTION, COMPANY, BODY OR ASSOCIATION"),
            "project_title": self.extract_field(combined_text, "PROJECT TITLE"),
            "implementation_period": self.extract_field_multiline(combined_text, "IMPLEMENTATION PERIOD (Defined as the duration where the RegTech solution is implemented): ")
        }

        # Save extracted data to JSON per application folder
        output_path = os.path.join(OUTPUT_DIR, f"{app_folder_name}_extracted.json")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(extracted, f, indent=4)
        logging.info(f"Extracted data saved to {output_path}")

    def run(self):
        for app_folder_name in os.listdir(INPUT_DIR):
            app_folder_path = os.path.join(INPUT_DIR, app_folder_name)
            if os.path.isdir(app_folder_path):
                file_paths = [
                    os.path.join(app_folder_path, f)
                    for f in os.listdir(app_folder_path)
                    if os.path.isfile(os.path.join(app_folder_path, f))
                ]
                self.process_application(app_folder_name, file_paths)

if __name__ == '__main__':
    processor = DocumentProcessor()
    processor.run()
