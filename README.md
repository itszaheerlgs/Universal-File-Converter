# 🔄 Universal File Converter

A modern, fast, and feature-rich desktop application built with Python and Tkinter that supports seamless cross-format conversions for Documents, Spreadsheets, Images, and specialized PDF tools. 

Designed with a sleek, responsive **Dracula-inspired dark theme** interface, multi-threading support to ensure zero UI freezes, and smart automatic output pathing.

---

## 🎨 Preview & UI Design

* **Left Panel:** Categorized list view with real-time search filtering.
* **Main Dashboard:** Drag-and-drop workflow feel with contextual configuration settings depending on the conversion type (e.g., rotation degrees, passwords, image formats).
* **Integrated Logging & Progress:** Live progress tracker and comprehensive logging panel built directly into the UI.

---

## ✨ Features & Supported Conversions

The application dynamically adjusts options and supports a versatile catalog of conversions:

### 📄 Documents
* **PDF ↔ Word** (`.pdf` to `.docx` and vice versa with automatic LibreOffice detection fallback)
* **PDF ↔ Text** (`.pdf` to `.txt` and vice versa)
* **Markdown → HTML** (complete with responsive styling rules, tables, and fenced code blocks)
* **HTML → Text** (strips markup cleanly)
* **Word → Text** (`.docx` to `.txt`)

### 📊 Spreadsheets
* **Excel ↔ CSV** (`.xlsx`/`.xls` to `.csv` and vice versa)
* **JSON ↔ CSV** (handles nested dictionary arrays or single JSON configurations seamlessly)

### 🖼 Images
* **Image → PDF** (combines individual image fields to full scalable page PDFs)
* **PDF → Images (PNG)** (extracts every page to clean `_page_001.png` file blocks)
* **Format Sweeper** (`.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.webp`, `.gif` cross-compatibility)

### 🔧 PDF Tools
* **Merge PDFs:** Chain multiple files in custom selected sequences.
* **Split PDF:** Automatically breaks a massive PDF file out into separate pages.
* **Rotate Pages:** Custom rotation input parameters ($90^\circ$, $180^\circ$, $270^\circ$).
* **Encrypt PDF:** Secure output streams with customized runtime password matrices.

---

## ⚙️ Requirements & Dependencies

The backend utilizes robust libraries to maximize speed and document integrity. Install the required components using `pip`:

```bash
pip install openpyxl markdown2 pdfplumber pdf2docx pypdf pypdfium2 pdf2image Pillow python-docx reportlab
