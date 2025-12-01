# Word to Markdown Converter

Simple drag-and-drop tool to convert Word documents (.doc/.docx) to Markdown with automatic image extraction.

## What it does

- Converts Word documents to clean Markdown format
- Automatically extracts and links images to a `media/` folder
- Supports page range selection (requires Microsoft Word on Windows)
- Works via drag-and-drop or GUI
- Auto-installs dependencies if missing

## Usage

**Drag & Drop:**
Drop a .docx file onto `word_to_markdown.pyw` → Get a .md file + images folder

**GUI Mode:**
Run `word_to_markdown.pyw` without arguments for a simple interface with:
- File browser
- Optional page range extraction (Windows + Word only)
- Visual conversion progress

## Features

- **Image handling:** Fixes MarkItDown's broken base64 placeholders, extracts actual images from docx
- **Page extraction:** Select specific page ranges (requires Microsoft Word on Windows)
- **.doc support:** Auto-converts old .doc files (requires Microsoft Word on Windows)
- **Smart output:** Places .md file next to original or in configured folder
- **Clean results:** Removes empty media folders automatically

## Requirements

- Python 3.8+
- `markitdown` (auto-installed)
- `pywin32` (auto-installed on Windows, optional for page extraction)
- Microsoft Word (optional, for .doc files and page range features)

## Configuration

Edit these variables in the script:
```python
USE_FIXED_OUTPUT_FOLDER = False  # True = all output to one folder
FIXED_OUTPUT_FOLDER = r"C:\Markdown_Output"
MEDIA_FOLDER_NAME = "media"
AUTO_INSTALL_DEPS = True
```
