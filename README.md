# PDF to DOCX Converter

A simple Python script that converts PDF files to Microsoft Word (DOCX) format using OCR-free text extraction.

## Features

- Converts PDF files to DOCX format
- Preserves text content from PDFs
- Easy-to-use command-line interface
- No OCR required - extracts native text from PDFs

## Prerequisites

- Python 3.7 or higher
- Virtual environment (automatically set up)

## Installation

1. Clone or download this project to your computer
2. Open PowerShell and navigate to the project directory:
   ```powershell
   cd "path\to\your\pdfToDocx"
   ```
3. The required packages (`python-docx` and `pdfplumber`) are already installed in the virtual environment.

## How to Use

### Available Conversion Options

This project now includes **three different converters** with varying levels of formatting preservation:

#### 1. **Basic Converter** (`pdfDocx.py`)
- ✅ Fast and reliable
- ✅ Handles quotes in file paths automatically  
- ⚠️ Plain text output (minimal formatting)
```powershell
.\.venv\Scripts\python.exe pdfDocx.py
```

#### 2. **Advanced Converter** (`advanced_converter.py`)
- ✅ Preserves tables and their structure
- ✅ Detects and formats headers
- ✅ Maintains bullet points and lists
- ✅ Preserves paragraph spacing
```powershell
.\.venv\Scripts\python.exe advanced_converter.py
```

#### 3. **Ultra Converter** (`ultra_converter.py`)
- ✅ Maximum formatting preservation
- ✅ Font size detection and preservation
- ✅ Advanced text classification
- ✅ Contact information formatting
- ✅ Section header emphasis
- ⚠️ Requires additional packages (auto-installed)
```powershell
.\.venv\Scripts\python.exe ultra_converter.py
```

### Running the Scripts

1. Open PowerShell and navigate to the project directory
2. Choose your preferred converter and run it:
3. The script will prompt you: `Enter the path to the PDF file:`

### Providing PDF File Paths

You have several options for specifying the PDF file:

#### Option 1: Full File Path
Type the complete path to your PDF file:
```
C:\Users\YourName\Documents\MyFile.pdf
```

#### Option 2: Relative Path
If your PDF is in the same folder as the script:
```
MyFile.pdf
```

#### Option 3: Drag and Drop (Recommended)
1. When prompted for the file path, don't type anything yet
2. Open File Explorer and locate your PDF file
3. Drag the PDF file from File Explorer into the PowerShell window
4. The full path will automatically appear
5. Press Enter to confirm

#### Option 4: Browse and Copy Path
1. Right-click on your PDF file in File Explorer
2. Hold Shift and right-click → "Copy as path"
3. Paste the path into the terminal (Ctrl+V)
4. Press Enter

### Example Usage

```
Enter the path to the PDF file: C:\Users\YourName\Documents\MyDocument.pdf
Document extracted and saved.
```

## Output Location

The converted DOCX files will be saved in the **same directory as the script** with different filenames based on the converter used:

| Converter | Output Filename | Features |
|-----------|----------------|----------|
| Basic | `PDF-Enhanced-output.docx` | Plain text with basic structure |
| Advanced | `PDF-Advanced-Formatted.docx` | Tables, headers, bullets, spacing |
| Ultra | `PDF-Ultra-Formatted.docx` | Maximum formatting preservation |

**Full path:** `[Your project directory]\[Output filename]`

⚠️ **Note:** Each time you run a converter, it will overwrite its previous output file.

## Troubleshooting

### Common Issues

1. **File not found error:**
   - Make sure the PDF file path is correct
   - Check that the file exists and you have permission to read it
   - Use quotes around paths with spaces: `"C:\Path with spaces\file.pdf"`

2. **Permission errors:**
   - Make sure you have read access to the PDF file
   - Ensure the script directory is writable for the output file

3. **Empty output:**
   - Some PDFs may not contain extractable text (scanned images)
   - Try with a different PDF that contains selectable text

### File Path Tips

- Use forward slashes `/` or double backslashes `\\` in paths
- Single backslashes `\` may cause issues
- Enclose paths with spaces in quotes: `"My File.pdf"`

## Technical Details

### Dependencies
- **pdfplumber**: For extracting text from PDF files
- **python-docx**: For creating Microsoft Word documents

### What the Script Does
1. Prompts user for PDF file path
2. Opens and reads the PDF file page by page
3. Extracts text content from each page
4. Creates a new Word document
5. Adds each page's text as a paragraph
6. Saves the document as "PDF-Plumber-output.docx"

## File Structure

```
pdfToDocx/
├── .venv/                          # Virtual environment
├── pdfDocx.py                      # Main conversion script
├── README.md                       # This file
└── PDF-Plumber-output.docx         # Output file (created after running)
```

## Support

If you encounter any issues:
1. Make sure you're using the correct Python command with the virtual environment
2. Verify the PDF file exists and is readable
3. Check that the PDF contains extractable text (not just images)
4. Ensure you have write permissions in the script directory