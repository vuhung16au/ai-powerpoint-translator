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
├── logs/                   # Directory containing translation logs
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

# Program Arguments

The slide translation tool supports various command-line arguments to customize its behavior:

## Basic Arguments

- `--tone`, `-t`: Specify the desired tone for translation
  - Options: formal, informal, technical, paper
  - Default: formal
  - Example: `python slide-tran.py --tone technical`

- `--version`, `-v`: Show version information and exit
  - Example: `python slide-tran.py --version`

- `--api_key`, `-k`: Provide the API key for the translation service
  - Example: `python slide-tran.py --api_key YOUR_API_KEY`

## Language Arguments

- `--source_language`, `--source-language`, `-sl`: Specify the source language
  - Default: Vietnamese
  - Example: `python slide-tran.py --source_language English`

- `--target_language`, `--target-language`, `-tl`: Specify the target language
  - Default: Japanese
  - Example: `python slide-tran.py --target_language Chinese`

- `--supported-languages`, `-sll`: List all supported languages and exit
  - Example: `python slide-tran.py --supported-languages`

## Directory Arguments

- `--input_dir`, `-i`: Specify the input directory containing PowerPoint files
  - Default: input
  - Example: `python slide-tran.py --input_dir my_slides`

- `--output_dir`, `-o`: Specify the output directory for translated files
  - Default: output
  - Example: `python slide-tran.py --output_dir translated_slides`

## Model Arguments

- `--llm_model`, `-m`: Specify the LLM model to use for translation
  - Default: gemini-2.0-flash-lite
  - Supported models: gemini-2.0-flash-lite, gpt-3.5-turbo, gpt-4, deepseek-chat
  - Example: `python slide-tran.py --llm_model gpt-4`

## RAG (Retrieval-Augmented Generation) Arguments

- `--rag`, `-rag`: Enable RAG for improved translation quality with domain-specific knowledge
  - Example: `python slide-tran.py --rag`

- `--rag-file`, `-rf`: Specify the file containing RAG context or examples
  - Default: rag_context.txt
  - Example: `python slide-tran.py --rag --rag-file my_context.txt`

## Examples

Translate PowerPoint from English to Japanese with formal tone using GPT-4:
```bash
python slide-tran.py --source_language English --target_language Japanese --tone formal --llm_model gpt-4
```

Translate PowerPoint from Vietnamese to Chinese using custom directories:
```bash
python slide-tran.py --source_language Vietnamese --target_language Chinese --input_dir source_slides --output_dir chinese_slides
```

Use RAG with custom context file for technical translations:
```bash
python slide-tran.py --source_language English --target_language Japanese --tone technical --rag --rag-file technical_terms.txt
```

# Supported Languages & Fallback Languages

## Officially Supported Languages

The tool officially supports the following 28 languages:

- Arabic
- Chinese
- Czech
- Danish
- Dutch
- English
- Finnish
- French
- German
- Greek
- Hebrew
- Hindi
- Hungarian
- Indonesian
- Italian
- Japanese
- Korean
- Norwegian
- Polish
- Portuguese
- Romanian
- Russian
- Spanish
- Swedish
- Thai
- Turkish
- Ukrainian
- Vietnamese

## Language Codes

The tool uses ISO 639-1 language codes (two-letter codes) for file naming and internal processing. These codes are loaded from the `lang-code.json` file, which contains mappings for over 180 languages.

For example:
- English: en
- Japanese: ja
- Chinese: zh
- Vietnamese: vi

## Fallback Mechanism

If the `lang-code.json` file cannot be loaded, the tool falls back to a limited set of hardcoded language codes:

```
Chinese: zh
English: en
Vietnamese: vi
Japanese: ja
```

## Adding Support for Additional Languages

While the tool officially supports 28 languages, you can attempt to use other languages listed in the `lang-code.json` file. The translation quality may vary based on the LLM model's capabilities with those languages.

To optimize translation for a specific language:

1. Ensure the language is included in the `lang-code.json` file
2. Create a custom prompt in your `Translation-Prompt.md` file that specifically addresses translation nuances for that language pair
3. Consider using RAG with domain-specific terminology when translating technical content

## Output File Naming

Translated files are saved with the language code appended to the filename:

```
original_filename_[language-code].pptx
```

For example, if you translate "presentation.pptx" to Japanese, the output file will be "presentation_ja.pptx".

# Known Issues

See [issues]

# TODOs

- Translate tables
- Translate images (?)
- Translate charts
- Translate shapes
- Translate notes
- Translate comments
- Translate SmartArt