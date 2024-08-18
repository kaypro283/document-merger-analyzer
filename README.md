# Document Merger, Analyzer, and Word Counter

Automate merging of DOC, DOCX, and PDF files with word frequency analysis. Streamlines document consolidation for large-scale projects.

## Features

- Converts DOC and DOCX files to PDF format
- Merges all PDF files into a single document
- Converts the merged PDF back to DOCX format
- Performs word frequency analysis on the final document
- Generates a detailed audit log of all operations

## Requirements

- Python 3.x
- Required Python packages:
  - PyPDF2
  - pdf2docx
  - win32com
  - tqdm

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/document-merger-analyzer.git
   ```
2. Navigate to the project directory:
   ```bash
   cd document-merger-analyzer
   ```
3. Install the required packages:
   ```bash
   pip install PyPDF2 pdf2docx pywin32 tqdm
   ```

## Usage

1. Run the script:
   ```bash
   python document_processor.py
   ```
2. Follow the prompts to:
   - Specify the input folder containing your documents
   - Name the output DOCX file
   - Enter words for frequency analysis

Example interaction:

```
Enter the path to the folder containing the files: C:\Users\YourName\Documents\InputFolder
Enter the name for the final output DOCX file (e.g., final_document.docx): merged_output.docx
Enter words to search for (one per line). Press Enter on a blank line to finish:
important
critical
urgent
```

The script will process the files and save the merged document and audit log in your Documents folder.

## Note

This script was created to handle a specific project involving merging hundreds of DOC files with some PDF files mixed in. It may require modifications for different use cases.

## Author

Christopher D. van der Kaay, Ph.D.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
