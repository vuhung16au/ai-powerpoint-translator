# Slide Translation Tool

A tool for automatically translating PowerPoint content using Gemini API. By default, it translates from Vietnamese to Japanese, but you can customize the source and target languages.

## System Requirements

- Python 3.8 or higher
- pip (Python package manager)

## Installation

1. Clone this repository to your local machine

2. Setup Python Virtual Environment and Install Required Libraries:

(On Mac/Linux)
```bash
python3 -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt
```

3. Set Gemini key

```bash
export GEMINI_API_KEY=<Your-API-Key>
```
(or you can define `GEMINI_API_KEY` in your `.env` file )

```bash
cat .env
GEMINI_API_KEY=<Your-API-Key>
```

You can get your Gemini API key from [Google AI website](https://ai.google.dev/gemini-api/docs/api-key). 

Note: Though being named Gemini API key, the code should be able to handle other keys such as DeepSeek, OpenAI API keys as well (TODO: with a little modification).

4. Create a `Translation-Prompt.md` file in the root directory with translation template content. You can customize the source and target languages by modifying the prompt content. For example:
   - For Vietnamese to Japanese: "You are a professional translator from Vietnamese to Japanese..."
   - For English to Japanese: "You are a professional translator from English to Japanese..."
   - For any other language pair: "You are a professional translator from [Source Language] to [Target Language]..."

TODO: Support multiple languages.

5. Prepare input directory and PowerPoint files:

   ```bash
   # Create input directory
   mkdir input
   
   # Copy PowerPoint file to be translated into input directory
   # Example: copy example.pptx to input directory
   cp example.pptx input/
   ```

   Notes:
   - Only place .pptx files in the input directory
   - (Pro tip) Ensure PowerPoint file is not open in PowerPoint application
   - (Pro tip) If PowerPoint file is open, close it before copying to input directory
   - Multiple PowerPoint files can be placed in input directory simultaneously

## Usage


1. Ensure you have completed all installation steps above

2. Run the script:

```bash
python slide-tran.py
```

3. The script will:
   - Create folder `output` directory to store translated files if not existed
   - Create folder `logs` directory to store translation logs if not existed
   - Translate all .pptx files under `input` directory
   - Save translated files to `output` directory with original filename + "_`target-language`" suffix

## Directory Structure

```
.
├── input/                  # Directory containing PowerPoint files to be translated
├── output/                 # Directory containing translated files
├── logs/      # Directory containing translation logs
├── .env                    # File containing API key
├── Translation-Prompt.md   # File containing translation prompts
├── requirements.txt        # Required Python libraries
└── slide-tran.py           # Main script
```

## Notes

- Ensure PowerPoint files are not open in PowerPoint application when running the script
- Script will automatically create necessary directories if they don't exist
- Each run will create a new log file in `logs` directory
- If output file already exists, script will automatically add sequence number to filename
- If you encounter errors while running the script, check the log file in `logs` directory for error details 
- (Pro tips) Check `logs/*.csv` for original and translated strings

# TODO 

- Add more arguments to `slide-tran.py`
- Handle loggings
- Handle errors, exceptions
- Create unit tests 
- Handle more languages (source languages and target languages)
- Handle more LLM models and consider what languages it supports (DeepSeek, OpenAI)
- Delete unused libraries (dotenv?)