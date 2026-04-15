# 📄 File Converter

A cross-platform desktop application for converting and manipulating documents and images. Built with a **PySide6** GUI frontend and a **FastAPI** backend, the app supports a wide range of file format conversions — all through a clean, sidebar-driven interface.

---

## ✨ Features

| Conversion / Operation | Description |
|---|---|
| PDF → Word | Convert PDF files to editable `.docx` format |
| Word → PDF | Convert `.docx` files to PDF |
| PDF → Image | Render each PDF page as an image, or extract embedded images |
| Image → PDF | Combine one or more images (JPG, PNG, WEBP) into a single PDF with orientation and scale control |
| Excel → PDF | Convert `.xlsx`/`.xls` spreadsheets to PDF |
| PPT → PDF | Convert PowerPoint presentations to PDF |
| Merge PDFs | Combine multiple PDF files into one |
| Split PDF | Extract a page range from a PDF |
| Rotate PDF | Rotate all pages of a PDF by 90°, 180°, or 270° |
| Watermark PDF | Add a custom diagonal text watermark to a PDF |
| Extract PDF Text | Extract plain text content from a PDF |

---

## 🗂️ Project Structure

```
File-Converter/
├── app.py               # PySide6 desktop GUI — sidebar navigation & API calls
├── converter.py         # Core conversion logic (LibreOffice, PyMuPDF, Pillow, etc.)
├── main.py              # FastAPI server — exposes REST endpoints for each conversion
├── ConverterSetup.iss   # Inno Setup script for building a Windows installer
└── File_Converter_Logo.ico
```

---

## 🏗️ Architecture

The application follows a client-server model:

- **Frontend (`app.py`)** — A PySide6 desktop app that provides the UI. When the user triggers a conversion, it sends a `multipart/form-data` POST request to the backend API.
- **Backend (`main.py`)** — A FastAPI server that receives files, delegates to `converter.py`, and returns the converted file as a downloadable response.
- **Converter (`converter.py`)** — The conversion engine. Uses LibreOffice (via subprocess), PyMuPDF, `pdf2docx`, Pillow, ReportLab, `pdfkit`, and `python-pptx` to handle each format.

The desktop client points to the backend at `http://43.204.22.70:8000` (AWS EC2). Converted files are saved directly to the user's `~/Downloads` folder.

---

## 🚀 Getting Started

### Prerequisites

- Python 3.9+
- [LibreOffice](https://www.libreoffice.org/download/download/) (must be installed and accessible via `soffice` in PATH)
- [wkhtmltopdf](https://wkhtmltopdf.org/downloads.html) (required by `pdfkit` for Excel → PDF)
- [Poppler](https://poppler.freedesktop.org/) (required by `pdf2image`)

### Installation

```bash
git clone https://github.com/Blackdevil-com/File-Converter.git
cd File-Converter
pip install -r requirements.txt
```

> **Note:** A `requirements.txt` is not yet included. Install the following packages manually:
>
> ```
> fastapi uvicorn pyside6 requests pymupdf pdf2docx pdf2image
> Pillow reportlab pdfkit python-pptx xlsx2html
> ```

### Running the Backend

```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

### Running the Desktop App

```bash
python app.py
```

> The app will connect to the backend. If you're running locally, update the `url_base` in `app.py` to `http://127.0.0.1:8000/convert/`.

---

## 🖥️ Building a Windows Installer

The project includes an [Inno Setup](https://jrsoftware.org/isinfo.php) script (`ConverterSetup.iss`) for packaging the app as a Windows `.exe` installer.

1. Install Inno Setup.
2. Open `ConverterSetup.iss` in the Inno Setup compiler.
3. Build to generate the installer.

---

## 🛠️ Tech Stack

- **GUI** — PySide6
- **API** — FastAPI
- **PDF Processing** — PyMuPDF (`fitz`), `pdf2docx`, `pdf2image`, ReportLab
- **Image Processing** — Pillow
- **Office Conversion** — LibreOffice (headless), `python-pptx`, `xlsx2html`, `pdfkit`
- **HTTP Client** — `requests`
- **Deployment** — AWS EC2

---

## 📌 Notes

- Converted files are automatically saved to the user's **Downloads** folder with unique filenames to avoid overwriting existing files.
- The PDF → Image mode supports two sub-modes: **page** (render each page) and **extract** (pull embedded images from the PDF).
- Image → PDF conversion supports **portrait/landscape** orientation and **full/medium/small** scale options.

---

## 📄 License

This project is open source. Feel free to use, modify, and distribute it.
