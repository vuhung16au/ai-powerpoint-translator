# Add style/tone in response 

- Add style/tone to the prompt/script arguments: Users can specify the desired style or tone for the translation, such as formal, informal, technical, etc. This can be done by adding a new argument to the `slide-tran.py` script and modifying the translation prompt accordingly.

Arguments to be added:
--tone or (-t): Specify the desired tone for the translation (e.g., formal, informal, technical, paper). 

Tone will be added to {tone} in the prompt from `Translation-Prompt.md`.

- Add more arguments to `slide-tran.py`

--help or (-h): Show help message and exit.
--version or (-v): Show version information and exit.
--api_key or (-k): Specify the API key for the translation service. Default is `os.getenv("API_KEY")`.

--source_language or (-sl): Specify the source language of the text to be translated (e.g., Vietnamese, English, etc.). 
--target_language or (-tl): Specify the target language for the translation (e.g., Japanese, English, etc.).
--input_dir or (-i): Specify the input directory containing PowerPoint files to be translated. Default is `input/`.
--output_dir or (-o): Specify the output directory where translated files will be saved. Default is `output/`.

# Response tone:

The response tone can be specified in the prompt. For example, you can add a line in the prompt like:

--tone="formal" will be set to {tone} in the prompt from `Translation-Prompt.md`.

# Supported languages:
Arugments to be added:
--supported-languages or (-sll): List of supported languages for translation. 

--source_language="value" will be set to {source_language} in the prompt from `Translation-Prompt.md`.
---target_language="value" will be set to {target_language} in the prompt from `Translation-Prompt.md`.

# Handle more LLM models and consider what languages it supports (DeepSeek, OpenAI)

Arguments to be added:
--llm_model or (-m): Specify the LLM model to be used for translation (e.g., DeepSeek, OpenAI, etc.). Default is `DeepSeek`.

# Use RAG to improve translation quality
- Use RAG (Retrieval-Augmented Generation) to improve translation quality by retrieving relevant context or examples from a database or knowledge base before generating the translation. This can be done by integrating a retrieval system into the translation process.

--rag or (-rag): Enable RAG for improved translation quality. Default is `False`.
--rag-file or (-rf): Specify the file containing the RAG context or examples. Default is `rag_context.txt`.

Example:

--rag=True --rag-file="rag_context.txt"

Our application will use RAG to improve translation quality by retrieving relevant context or examples from the specified file before generating the translation.


# Handle errors, exceptions
- Implement error handling and exception management in the script to ensure that any issues encountered during the translation process are logged and handled gracefully. This can include handling file not found errors, API errors, and other unexpected issues.
- Implement logging to capture errors and exceptions in a structured manner, allowing for easier debugging and troubleshooting.

# Add langugage code at the end of the file name
- Add language code at the end of the translated file name to indicate the target language of the translation. For example, if the original file is `example.pptx` and the target language is Japanese, the translated file should be named `example_ja.pptx`.
- For Chinese, use "zh" for simplified Chinese and "zh_tw" for traditional Chinese.
- For Japanese, use "ja" for Japanese.
- For Korean, use "ko" for Korean.
- For Vietnamese, use "vi" for Vietnamese.
- For English, use "en" for English.
- ...

# TODO 

- Handle more file formats (e.g., .docx, .txt)

