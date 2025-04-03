import os
import glob
from pptx import Presentation
from openai import OpenAI
from dotenv import load_dotenv
import logging
import time
from datetime import datetime
import httpx
import sys

# Set up logging first, before any other imports
def setup_logging():
    # Create logs directory if it doesn't exist
    log_dir = 'SlideTranslateLog'
    os.makedirs(log_dir, exist_ok=True)
    
    # Create a timestamp for the log file
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = os.path.join(log_dir, f'translation_{timestamp}.log')
    
    # Configure logging to only file
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8')
        ]
    )
    return log_file

# Initialize logging first
log_file = setup_logging()
logging.info("Logging system initialized")

# Load environment variables
load_dotenv()


# Initialize OpenAI client
# Tạo một HTTP client tùy chỉnh, ở đây chúng ta không cấu hình proxy
# Nếu bạn CẦN dùng proxy, bạn phải cấu hình nó đúng cách tại đây.
# Ví dụ: proxies = {"http://": os.getenv("HTTP_PROXY"), "https://": os.getenv("HTTPS_PROXY")}
custom_http_client = httpx.Client() # Explicitly disable proxies if not needed

client = OpenAI(
    api_key=os.getenv('GEMINI_API_KEY'),
    base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
    http_client=custom_http_client # Truyền HTTP client tùy chỉnh vào
)


# Load translation prompt
with open('prompt.txt', 'r', encoding='utf-8') as f:
    PROMPT_TEMPLATE = f.read()

def batch_texts(texts, batch_size=30):
    """Group texts into batches for translation."""
    return [texts[i:i + batch_size] for i in range(0, len(texts), batch_size)]

def translate_batch(texts):
    """Translate a batch of texts from Vietnamese to Japanese."""
    if not texts:
        return []
    
    prompt = PROMPT_TEMPLATE.format(texts="\n---\n".join(texts))
    
    # Log the request
    logging.info("=== Translation Request ===")
    for idx, text in enumerate(texts):
        logging.info(f"Text {idx + 1}: {text}")
    logging.info("=== End Request ===\n")
    
    try:
        response = client.chat.completions.create(
            model="gemini-2.0-flash-lite",
            n=1,
            messages=[
                {"role": "system", "content": "You are a professional translator from Vietnamese to Japanese."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            stream=False
        )
        
        # Log the response
        logging.info("=== Translation Response ===")
        logging.info(f"Raw response: {response.choices[0].message.content}")
        logging.info("=== End Response ===\n")
        
        # Parse the response to get translations
        translations = response.choices[0].message.content.strip().split("\n---\n")
        
        # Log parsed translations
        logging.info("=== Parsed Translations ===")
        for idx, (text, trans) in enumerate(zip(texts, translations)):
            logging.info(f"Text {idx + 1}:")
            logging.info(f"Original: {text}")
            logging.info(f"Translated: {trans}")
            logging.info("---")
        logging.info("=== End Parsed Translations ===\n")
        
        # Ensure we have the same number of translations as input texts
        if len(translations) != len(texts):
            logging.warning(f"Received {len(translations)} translations for {len(texts)} texts")
            # Pad or trim translations to match input count
            if len(translations) < len(texts):
                translations.extend([""] * (len(texts) - len(translations)))
                logging.warning("Added empty translations to match input count")
            else:
                translations = translations[:len(texts)]
                logging.warning("Trimmed excess translations")
        
        return translations
        
    except Exception as e:
        logging.error(f"Error during translation: {str(e)}")
        raise

def save_presentation(prs, original_filename):
    """Save presentation with error handling and unique filename."""
    # Create output directory if it doesn't exist
    output_dir = 'output'
    os.makedirs(output_dir, exist_ok=True)
    
    # Generate base output filename
    base_name = os.path.basename(original_filename)
    name_without_ext = os.path.splitext(base_name)[0]
    
    # Try to save with different names if file exists or is locked
    counter = 1
    while True:
        if counter == 1:
            output_filename = os.path.join(output_dir, f"{name_without_ext}_ja.pptx")
        else:
            output_filename = os.path.join(output_dir, f"{name_without_ext}_ja_{counter}.pptx")
        
        try:
            prs.save(output_filename)
            logging.info(f"Successfully saved presentation to {output_filename}")
            return output_filename
        except PermissionError:
            logging.warning(f"Permission denied when saving to {output_filename}. File might be open in PowerPoint.")
            logging.info("Please close the file in PowerPoint if it's open.")
            counter += 1
            if counter > 5:  # Limit number of attempts
                raise Exception(f"Failed to save presentation after {counter-1} attempts. Please ensure the file is not open in PowerPoint.")
        except Exception as e:
            logging.error(f"Error saving presentation: {str(e)}")
            raise

def extract_table_texts(shape):
    """Extract texts from a table shape."""
    texts = []
    locations = []
    
    if not hasattr(shape, "table"):
        return texts, locations
        
    for row_idx, row in enumerate(shape.table.rows):
        for cell_idx, cell in enumerate(row.cells):
            if cell.text.strip():
                texts.append(cell.text.strip())
                locations.append((row_idx, cell_idx))
    
    return texts, locations

def process_presentation(input_file):
    """Process a PowerPoint presentation, translating text from Vietnamese to Japanese."""
    logging.info(f"Processing {input_file}")
    
    try:
        prs = Presentation(input_file)
        all_texts = []
        text_locations = []
        
        # Extract all texts that need translation
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                # Handle regular text shapes
                if hasattr(shape, "text") and shape.text.strip():
                    all_texts.append(shape.text.strip())
                    text_locations.append(("shape", slide_idx, shape_idx))
                
                # Handle tables
                if hasattr(shape, "table"):
                    table_texts, table_locations = extract_table_texts(shape)
                    for text, (row_idx, cell_idx) in zip(table_texts, table_locations):
                        all_texts.append(text)
                        text_locations.append(("table", slide_idx, shape_idx, row_idx, cell_idx))
        
        if not all_texts:
            logging.info(f"No text found in {input_file}")
            return
        
        # Translate texts in batches
        translated_texts = []
        batches = batch_texts(all_texts)
        
        print(f"\nTranslating {os.path.basename(input_file)}:")
        for i, batch in enumerate(batches):
            progress = (i + 1) / len(batches) * 100
            sys.stdout.write(f"\rProgress: [{int(progress)}%] Batch {i+1}/{len(batches)}")
            sys.stdout.flush()
            
            logging.info(f"Translating batch {i+1}/{len(batches)} (size: {len(batch)} texts)")
            translations = translate_batch(batch)
            translated_texts.extend(translations)
            
            if i < len(batches) - 1:
                time.sleep(2)
        print("\nTranslation completed!")
        
        # Update presentation with translations
        for location, translated_text in zip(text_locations, translated_texts):
            if location[0] == "shape":
                _, slide_idx, shape_idx = location
                shape = prs.slides[slide_idx].shapes[shape_idx]
                if hasattr(shape, "text_frame"):
                    # Replace text while preserving formatting
                    for i, paragraph in enumerate(shape.text_frame.paragraphs):
                        if i == 0:
                            paragraph.text = translated_text
                        else:
                            paragraph.text = ""
            else:  # table
                _, slide_idx, shape_idx, row_idx, cell_idx = location
                shape = prs.slides[slide_idx].shapes[shape_idx]
                if hasattr(shape, "table"):
                    cell = shape.table.rows[row_idx].cells[cell_idx]
                    if hasattr(cell, "text_frame"):
                        for i, paragraph in enumerate(cell.text_frame.paragraphs):
                            if i == 0:
                                paragraph.text = translated_text
                            else:
                                paragraph.text = ""
        
        # Save translated presentation with error handling
        save_presentation(prs, input_file)
        
    except Exception as e:
        logging.error(f"Error processing presentation {input_file}: {str(e)}")
        raise

def main():
    # Setup logging
    log_file = setup_logging()
    logging.info(f"Translation log file: {log_file}")
    
    # Find all PPTX files in the input directory
    input_files = glob.glob('input/*.pptx')
    
    if not input_files:
        logging.warning("No PowerPoint files found in the input directory")
        return
    
    for input_file in input_files:
        logging.info(f"\n=== Processing file: {input_file} ===")
        process_presentation(input_file)
        logging.info(f"Completed translation of {input_file}")

if __name__ == "__main__":
    main()