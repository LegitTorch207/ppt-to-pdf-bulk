# PPT to PDF Bulk Converter

This project allows you to **convert multiple PowerPoint files (.ppt and .pptx) to PDF** at once.  
It uses a **PowerShell script** and a **batch file** for drag-and-drop convenience.

## Features

- Converts `.ppt` and `.pptx` files to PDF in bulk.
- Creates a `PDFs` folder inside the selected folder automatically.
- Skips files that are already converted.
- Shows conversion progress and a summary at the end.
- Works by **dragging a folder onto the batch file**.

## How to Use

1. Download or clone the repository.
2. Make sure you have **PowerPoint installed** on your computer.
3. Place the folder you want to convert somewhere accessible.
4. Drag and drop the folder onto `convert-ppt-to-pdf.bat`.
5. Wait for the script to finish. Converted PDFs will appear in the `PDFs` folder inside your folder.

## Files

- `convert-ppt-to-pdf.ps1` - The PowerShell script that does the conversion.
- `convert-ppt-to-pdf.bat` - Batch file for drag-and-drop functionality.

## Notes

- This script runs silently but requires **PowerPoint** installed on your system.
- Tested on Windows 10/11 with Microsoft PowerPoint.
