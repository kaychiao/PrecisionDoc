# PrecisionDoc - Medical Precision Document Processing Tool

This project processes medical guideline PDF files, especially treatment guidelines from CSCO (Chinese Society of Clinical Oncology). It can:

1. Process PDF files in a specified folder
2. Split PDF files into individual pages
3. Analyze each page using AI (OpenAI or Alibaba Cloud Qwen)
4. Extract precision medicine evidence related to drug efficacy
5. Save analysis results in JSON and Excel formats
6. Generate Word reports containing precision medicine evidence

## Installation

1. Clone this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Create a `.env` file (refer to `env.example`) and set API keys:

```
OPENAI_API_KEY=your_openai_api_key
OPENAI_BASE_URL=https://api.openai.com/v1
OPENAI_MODEL=gpt-4

QWEN_API_KEY=your_qwen_api_key
QWEN_BASES_URL=https://dashscope.aliyuncs.com/compatible-mode/v1
QWEN_TEXT_MODEL=qwen-max
QWEN_MULTIMODAL_MODEL=qwen-vl-max

LOG_LEVEL=INFO
```

### Dependencies

The project requires the following main dependencies:
- `PyMuPDF`: PDF processing
- `openai`: OpenAI API client
- `pandas` and `openpyxl`: Data processing and Excel file handling
- `python-docx`: Word document generation
- `python-dotenv`: Environment variable management
- `numpy`: Numerical operations
- `requests`: HTTP requests
- `tqdm`: Progress bars

All dependencies are listed in `requirements.txt`.

## Usage

Basic usage:

```bash
python main.py --folder /path/to/pdf/folder --api-key YOUR_API_KEY
```

Using Alibaba Cloud Qwen API instead of OpenAI:

```bash
python main.py --folder /path/to/pdf/folder --api-key YOUR_API_KEY --use-qwen
```

## Parameters

- `--folder`: Path to the folder containing PDF files (required)
- `--api-key`: API key for OpenAI or Qwen (if not provided, will be read from environment variables)
- `--use-qwen`: Use Qwen API instead of OpenAI (optional)

## Output

The program creates the following in the `./output` directory:

- `pages`: Contains split single-page PDF files
- `images`: (When using Qwen) Contains PDF page image files
- JSON files: Structured data with AI processing results
- Excel files: Flattened analysis results
- Word files: Extracted precision medicine evidence reports

## Features

- Automatic page type identification (content pages, table of contents, references)
- Extraction of precision medicine evidence, including:
  - Related genes and alterations
  - Disease names (Chinese and English)
  - Drug names (Chinese and English) and drug combinations
  - Evidence levels and response types
- Support for text-based (OpenAI and Qwen) and image-based (Qwen only) analysis
- Generation of structured JSON output and easy-to-read Excel spreadsheets
- Generation of formatted Word report documents

## Word Export Features

The Word export functionality includes several advanced formatting options:

- **Enhanced Table Layout**: 
  - Left side displays multiple rows of text fields (one field per row)
  - Right side shows images in a single vertically merged cell
  - Customizable table borders (can be shown or hidden)

- **Page Formatting**:
  - Automatic page numbers in "current/total" format (e.g., "3 / 10")
  - Support for both portrait and landscape orientations
  - Table continuation across pages for long evidence items
  - Clear separator lines between evidence items

- **Text Formatting**:
  - Support for multi-line text in evidence fields
  - Consistent font styling and paragraph formatting
  - Proper handling of Chinese and English text

- **Image Handling**:
  - Automatic resizing and centering of images
  - Fallback mechanisms for missing images
  - Support for various image formats

To customize Word export, you can modify the parameters in `export_evidence_to_word()`:

```python
WordUtils.export_evidence_to_word(
    excel_file="path/to/excel.xlsx",
    word_file="path/to/output.docx",
    output_folder="path/to/images",
    multi_line_text=True,  # Set to False for single-line text
    show_borders=True      # Set to False to hide table borders
)
```

## Notes

1. Ensure you have sufficient API call quota, as each PDF page will make one AI API call
2. Processing large PDF files may take a significant amount of time
3. Image processing features are only available when using Qwen API
4. Make sure all necessary dependencies are installed in your Python environment
