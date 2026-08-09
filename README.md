# 📑 PDF Converter

![Batch Script](https://img.shields.io/badge/script-batch-blue.svg)
![Requires](https://img.shields.io/badge/requires-Microsoft%20Office-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows-0078D4.svg)
![License](https://img.shields.io/badge/RaneKun-Open--Use--License-yellow.svg)

A powerful Windows batch script that converts multiple file formats to PDF using Microsoft Office automation. Supports Word, PowerPoint, text files, and images with automatic page fitting and centering.

## ✨ Features

- 🔄 **Converts multiple formats** - DOCX, DOC, PPTX, PPT, TXT, PNG, JPG, JPEG
- 📁 **Batch processing** - converts all matching files in the current folder
- 🖼️ **Smart image placement** - auto-centers and fits images to page
- 🧹 **Two versions available** - keeps originals OR removes them after conversion
- 🚫 **No popups** - runs silently with minimal interruption

## 📦 Two Versions

| File | Behavior |
|------|----------|
| `PDF Converter.bat` | Keeps original files after conversion |
| `PDF Converter [remove original files].bat` | Deletes originals after successful conversion |

## 🚀 Quick Start

1. Ensure Microsoft Office (Word, PowerPoint) is installed

2. Place the script in the folder with your files

3. Double-click the script

4. Find your PDFs in the same folder!

## 📁 Supported Formats

| Format | Tool Used | Result |
|--------|-----------|--------|
| `.docx` / `.doc` | Microsoft Word | Perfect document conversion |
| `.pptx` / `.ppt` | Microsoft PowerPoint | Slide-to-PDF conversion |
| `.txt` | Microsoft Word | Text document to PDF |
| `.png` / `.jpg` / `.jpeg` | Microsoft Word | Centered, full-page images |

## ⚙️ How It Works

Each file type is handled with specific automation:

- **Word docs** → Opens, saves as PDF (Format 17), closes
- **PowerPoint** → Opens, saves as PDF (Format 32), closes  
- **Images** → Creates new Word doc, inserts image, centers it, saves as PDF

## ⚠️ Important Notes

### Before Running:
- ✅ Close **ALL** open Microsoft Office applications (Word, PowerPoint, etc.)
- ✅ Make sure files aren't open in another program
- ✅ Files must not be read-only

### For PowerPoint (Optional):
To avoid a popup when converting PPT files:

1. Go to **Bluetooth & devices > Printers and scanners**
2. Find **Microsoft Print to PDF**
3. Set it as **default printer**
