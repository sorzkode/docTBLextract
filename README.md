# docTBLextract - Word to Excel Table Converter

docTBLextract is a tool for converting tables from Microsoft Word documents to Excel worksheets. Each table in the Word document becomes a separate worksheet in the Excel file with automatic column width adjustment.

## Features

- **Table extraction** - Extract all tables from Word documents
- **Excel conversion** - Convert each table to a separate Excel worksheet
- **Auto-formatting** - Automatic column width adjustment
- **Random naming** - Each worksheet gets a unique random name
- **Progress tracking** - Real-time conversion progress
- **Error handling** - Clear error messages and status updates
- **GUI interface** - Easy-to-use tkinter interface

## Installation

### Option 1: Download Executable (No Python Required)
Download the latest `docTBLextract.exe` from the [Releases](https://github.com/sorzkode/doctblextract/releases) page and run directly.

### Option 2: Install with Python
**Install directly from GitHub:**
```bash
pip install git+https://github.com/sorzkode/doctblextract.git
```

**Or install locally:**
1. Download/clone the repository
2. Navigate to the project directory
3. Run:
```bash
pip install .
```

## Requirements

- **For executable**: None - runs on any Windows machine
- **For Python version**: Python 3.8+ with required packages:
  - tkinter (included with most Python installations)
  - python-docx
  - openpyxl

## Usage

**Executable:**
```bash
# Simply run the downloaded file
docTBLextract.exe
```

**Python installation:**
```bash
doctblextract
```

**Run directly from source:**
```bash
python docTBLextract.py
```

Once the application is running:
```
  1. Click "Browse" to select a Word document (.docx)
  2. The application will analyze and show table count
  3. Click "Convert to Excel" to choose save location
  4. Wait for conversion to complete
```

## How It Works

1. **Document Analysis** - Scans the Word document for tables
2. **Table Extraction** - Extracts text content from each table cell
3. **Excel Creation** - Creates a new Excel workbook
4. **Worksheet Generation** - Each table becomes a separate worksheet with a random name
5. **Formatting** - Auto-adjusts column widths for readability
6. **File Save** - Saves the Excel file to your chosen location

## Supported Formats

- **Input**: Microsoft Word documents (.docx)
- **Output**: Excel workbooks (.xlsx)

## Notes

- Only .docx files are supported (not legacy .doc files)
- Tables are converted as-is with basic text formatting
- Complex table features (merged cells, styling) may not be preserved
- Each table gets a randomly generated worksheet name
- Column widths are automatically adjusted up to 50 characters maximum

## License

MIT License