# Slide Translation Tool

A tool for automatically translating PowerPoint content using Gemini API. By default, it translates from Vietnamese to Japanese, but you can customize the source and target languages.

## System Requirements

- Python 3.8 or higher
- pip (Python package manager)

## Installation

1. Clone this repository to your local machine

2. Install required libraries:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the root directory and add your Gemini API key:
```
GEMINI_API_KEY=your_api_key_here
```

4. Create a `prompt.txt` file in the root directory with translation template content. You can customize the source and target languages by modifying the prompt content. For example:
   - For Vietnamese to Japanese: "You are a professional translator from Vietnamese to Japanese..."
   - For English to Japanese: "You are a professional translator from English to Japanese..."
   - For any other language pair: "You are a professional translator from [Source Language] to [Target Language]..."

5. Prepare input directory and PowerPoint files:
   ```bash
   # Create input directory
   mkdir input
   
   # Copy PowerPoint file to be translated into input directory
   # Example: copy example.pptx to input directory
   cp example.pptx input/
   ```
   
   Note:
   - Only place .pptx files in the input directory
   - Ensure PowerPoint file is not open in PowerPoint application
   - If PowerPoint file is open, close it before copying to input directory
   - Multiple PowerPoint files can be placed in input directory simultaneously

## Usage

1. Ensure you have completed all installation steps above

2. Run the script:
```bash
python slide-tran.py
```

3. The script will:
   - Automatically create an `output` directory to store translated files
   - Automatically create a `SlideTranslateLog` directory to store translation logs
   - Process all .pptx files in the `input` directory
   - Save translated files to `output` directory with original filename + "_ja" suffix

## Directory Structure

```
.
├── input/                  # Directory containing PowerPoint files to be translated
├── output/                 # Directory containing translated files
├── SlideTranslateLog/      # Directory containing translation logs
├── .env                    # File containing API key
├── prompt.txt             # File containing translation 
├── requirements.txt       # File containing required Python 
└── slide-tran.py         # Main script
```

## Notes

- Ensure PowerPoint file is not open in PowerPoint application when running the script
- Script will automatically create necessary directories if they don't exist
- Each run will create a new log file in SlideTranslateLog directory
- If output file already exists, script will automatically add sequence number to filename
- If you encounter errors while running the script, check the log file in SlideTranslateLog directory for error details 