import os
import sys
from pathlib import Path
import requests
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QListWidget, QFileDialog, QStackedWidget,
    QComboBox, QLineEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtGui import QIcon


def resource_path(relative_path):
    """Get absolute path to resource, works for dev and PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def get_unique_filepath(folder, filename):
    folder = Path(folder)
    file_path = folder / filename
    stem = file_path.stem
    suffix = file_path.suffix
    counter = 1
    while file_path.exists():
        file_path = folder / f"{stem} ({counter}){suffix}"
        counter += 1
    return str(file_path)


class ConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Converter")
        self.setWindowIcon(QIcon(resource_path("File_Converter_Logo.ico")))
        self.setGeometry(200, 100, 1200, 700)

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QHBoxLayout()
        main_widget.setLayout(main_layout)

        # ---------------- Left Sidebar ----------------
        self.sidebar = QListWidget()
        self.sidebar.addItems([
            "Home",
            "PDF → Word",
            "Word → PDF",
            "PDF → Image",
            "Image → PDF",
            "Excel → PDF",
            "PPT → PDF",
            "Merge PDFs",
            "Split PDF",
            "Rotate PDF",
            "Watermark PDF",
            "Extract PDF Text",
            # "OCR PDF"
        ])
        self.sidebar.setFixedWidth(220)
        self.sidebar.currentRowChanged.connect(self.display_page)
        main_layout.addWidget(self.sidebar)


        # ---------------- Main Panel ----------------
        self.stack = QStackedWidget()
        main_layout.addWidget(self.stack)

        # ---------------- File labels ----------------
        self.pdf_to_word_label = QLabel("No file selected")
        self.word_to_pdf_label = QLabel("No file selected")
        self.pdf_image_label = QLabel("No file selected")
        self.images_to_pdf_label = QLabel("No files selected")
        self.excel_to_pdf_label = QLabel("No file selected")
        self.ppt_to_pdf_label = QLabel("No file selected")
        self.merge_pdfs_label = QLabel("No files selected")
        self.split_pdf_label = QLabel("No file selected")
        self.rotate_pdf_label = QLabel("No file selected")
        self.watermark_pdf_label = QLabel("No files selected")
        self.extract_text_pdf_label = QLabel("No file selected")
        self.ocr_pdf_label = QLabel("No file selected")

        self.pages = {}
        self.create_pages()
        self.stack.setCurrentIndex(0)

    # ---------------- Page Creators ----------------
    def create_pages(self):
        home_page = QWidget()
        layout = QVBoxLayout()
        label = QLabel("Welcome to Universal Doc Converter!\nSelect an option from the left menu.")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)
        home_page.setLayout(layout)
        self.stack.addWidget(home_page)

        # Standard Conversions
        self.pages['pdf_to_word'] = self.create_file_page("Select PDF File", "Convert to Word", "pdf_to_word", self.pdf_to_word_label)
        self.pages['word_to_pdf'] = self.create_file_page("Select Word File", "Convert to PDF", "word_to_pdf", self.word_to_pdf_label)
        self.pages['pdf_image'] = self.create_pdf_image_page()
        self.pages['images_to_pdf'] = self.create_images_to_pdf_page()
        self.pages['excel_to_pdf'] = self.create_file_page("Select Excel File", "Convert Excel to PDF", "excel_to_pdf", self.excel_to_pdf_label)
        self.pages['ppt_to_pdf'] = self.create_file_page("Select PPT File", "Convert PPT to PDF", "ppt_to_pdf", self.ppt_to_pdf_label)

        # PDF Utilities
        self.pages['merge_pdfs'] = self.create_file_page("Select PDF Files to Merge", "Merge PDFs", "merge_pdfs", self.merge_pdfs_label)
        self.pages['split_pdf'] = self.create_file_page("Select PDF File to Split", "Split PDF", "split_pdf", self.split_pdf_label, extra_input=True, extra_label="Pages (e.g., 1-3,5)")
        self.pages['rotate_pdf'] = self.create_file_page("Select PDF File to Rotate", "Rotate PDF", "rotate_pdf", self.rotate_pdf_label, extra_input=True, extra_label="Angle (90,180,270)")
        self.pages['watermark_pdf'] = self.create_watermark_pdf_page()
        self.pages['extract_text_pdf'] = self.create_file_page("Select PDF File", "Extract Text", "extract_text_pdf", self.extract_text_pdf_label)
        self.pages['ocr_pdf'] = self.create_file_page("Select PDF File", "OCR PDF", "ocr_pdf", self.ocr_pdf_label)

    def create_file_page(self, select_text, convert_text, mode, label_widget, extra_input=False, extra_label=""):
        page = QWidget()
        layout = QVBoxLayout()
        file_btn = QPushButton(select_text)
        file_btn.clicked.connect(lambda: self.select_file(mode))
        convert_btn = QPushButton(convert_text)
        convert_btn.clicked.connect(lambda: self.convert_file(mode))
        layout.addWidget(file_btn)
        layout.addWidget(label_widget)
        if extra_input:
            input_label = QLabel(extra_label)
            input_field = QLineEdit()
            setattr(self, f"{mode}_input", input_field)
            layout.addWidget(input_label)
            layout.addWidget(input_field)
        layout.addWidget(convert_btn)
        page.setLayout(layout)
        self.stack.addWidget(page)
        return page

    def create_multi_file_page(self, select_text, convert_text, mode, label_widget):
        page = QWidget()
        layout = QVBoxLayout()
        file_btn = QPushButton(select_text)
        file_btn.clicked.connect(lambda: self.select_file(mode))
        convert_btn = QPushButton(convert_text)
        convert_btn.clicked.connect(lambda: self.convert_file(mode))
        layout.addWidget(file_btn)
        layout.addWidget(label_widget)
        layout.addWidget(convert_btn)
        page.setLayout(layout)
        self.stack.addWidget(page)
        return page

    def create_pdf_image_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        file_btn = QPushButton("Select PDF File")
        file_btn.clicked.connect(lambda: self.select_file("pdf_image"))
        mode_combo = QComboBox()
        mode_combo.addItems(["page", "extract"])
        convert_btn = QPushButton("Convert PDF to Images")
        convert_btn.clicked.connect(lambda: self.convert_file("pdf_image"))
        layout.addWidget(file_btn)
        layout.addWidget(self.pdf_image_label)
        layout.addWidget(QLabel("Mode:"))
        layout.addWidget(mode_combo)
        layout.addWidget(convert_btn)
        page.setLayout(layout)
        self.pdf_image_mode_combo = mode_combo
        self.stack.addWidget(page)
        return page

    def create_images_to_pdf_page(self):
        page = QWidget()
        layout = QVBoxLayout()
        file_btn = QPushButton("Select Image Files")
        file_btn.clicked.connect(lambda: self.select_file("images_to_pdf"))
        orientation_combo = QComboBox()
        orientation_combo.addItems(["portrait", "landscape"])
        scale_combo = QComboBox()
        scale_combo.addItems(["full", "medium", "small"])
        convert_btn = QPushButton("Convert Images to PDF")
        convert_btn.clicked.connect(lambda: self.convert_file("images_to_pdf"))
        layout.addWidget(file_btn)
        layout.addWidget(self.images_to_pdf_label)
        layout.addWidget(QLabel("Orientation:"))
        layout.addWidget(orientation_combo)
        layout.addWidget(QLabel("Scale:"))
        layout.addWidget(scale_combo)
        layout.addWidget(convert_btn)
        page.setLayout(layout)
        self.images_to_pdf_orientation_combo = orientation_combo
        self.images_to_pdf_scale_combo = scale_combo
        self.stack.addWidget(page)
        return page

    def create_watermark_pdf_page(self):
        page = QWidget()
        layout = QVBoxLayout()

        # 1️⃣ File selection
        file_btn = QPushButton("Select PDF File")
        file_btn.clicked.connect(lambda: self.select_file("watermark_pdf"))

        # 2️⃣ Watermark text input
        watermark_label = QLabel("Enter Watermark Text:")
        watermark_input = QLineEdit()
        setattr(self, "watermark_pdf_input", watermark_input)  # store input for later

        # 3️⃣ Convert button
        convert_btn = QPushButton("Add Watermark")
        convert_btn.clicked.connect(lambda: self.convert_file("watermark_pdf"))

        # 4️⃣ Status label
        layout.addWidget(file_btn)
        layout.addWidget(self.watermark_pdf_label)  # shows selected file / result
        layout.addWidget(watermark_label)
        layout.addWidget(watermark_input)
        layout.addWidget(convert_btn)

        page.setLayout(layout)
        self.stack.addWidget(page)
        return page
    # ---------------- UI Logic ----------------
    def display_page(self, index):
        self.stack.setCurrentIndex(index)

    def select_file(self, file_type):
        # Map file_type to label attribute and file filters
        file_filters = {
            "pdf_to_word": "PDF Files (*.pdf)",
            "word_to_pdf": "Word Files (*.docx)",
            "pdf_image": "PDF Files (*.pdf)",
            "images_to_pdf": "Images (*.jpg *.jpeg *.png *.webp)",
            "excel_to_pdf": "Excel Files (*.xlsx *.xls)",
            "ppt_to_pdf": "PPT Files (*.ppt *.pptx)",
            "merge_pdfs": "PDF Files (*.pdf)",
            "split_pdf": "PDF Files (*.pdf)",
            "rotate_pdf": "PDF Files (*.pdf)",
            "watermark_pdf": "PDF Files (*.pdf)",
            "extract_text_pdf": "PDF Files (*.pdf)",
            "ocr_pdf": "PDF Files (*.pdf)"
        }

        multi_select = file_type in ["images_to_pdf", "merge_pdfs", "watermark_pdf"]

        if multi_select:
            files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", file_filters[file_type])
            if files:
                getattr(self, f"{file_type}_label").setText(", ".join(files))
        else:
            file, _ = QFileDialog.getOpenFileName(self, "Select File", "", file_filters[file_type])
            if file:
                getattr(self, f"{file_type}_label").setText(file)

    # ---------------- Conversion / PDF Operations ----------------
    def convert_file(self, mode):
        downloads = str(Path.home() / "Downloads")
        input_attr = getattr(self, f"{mode}_label", None)
        if not input_attr:
            return

        input_file_text = input_attr.text()
        if not input_file_text or "No file" in input_file_text:
            input_attr.setText("File not selected!")
            return

        input_files = input_file_text.split(", ") if "," in input_file_text else [input_file_text]
        url_base = "http://43.204.22.70:8000/convert/"

        try:
            # ---------------- PDF → Word ----------------
            if mode == "pdf_to_word":
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}pdf-to-word", files={"file": f})

                # if r.status_code == 200:
                #     # Save file
                #     cd = r.headers.get("Content-Disposition", "")
                #     if "filename=" in cd:
                #         filename = cd.split("filename=")[-1].strip('"')
                #     else:
                #         filename = Path(input_files[0]).stem + ".docx"
                #     out_file = get_unique_filepath(downloads, filename)
                #     with open(out_file, "wb") as wf:
                #         wf.write(r.content)
                #     input_attr.setText(f"Saved: {out_file}")
                # else:
                #     try:
                #         # If backend returned JSON error
                #         input_attr.setText(str(r.json()))
                #     except:
                #         input_attr.setText(f"Error: {r.text}")

            # ---------------- Word → PDF ----------------
            elif mode == "word_to_pdf":
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}word-to-pdf", files={"file": f})

            # ---------------- PDF → Image ----------------
            elif mode == "pdf_image":
                mode_value = self.pdf_image_mode_combo.currentText()
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}pdf-image",
                                      files={"file": f},
                                      data={"mode": mode_value})

            # ---------------- Images → PDF ----------------
            elif mode == "images_to_pdf":
                orientation = self.images_to_pdf_orientation_combo.currentText()
                scale = self.images_to_pdf_scale_combo.currentText()
                files_payload = [("files", (Path(f).name, open(f, "rb"))) for f in input_files]
                r = requests.post(f"{url_base}images-to-pdf",
                                  files=files_payload,
                                  data={"orientation": orientation, "scale": scale})

            # ---------------- Excel → PDF ----------------
            elif mode == "excel_to_pdf":
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}excel-to-pdf", files={"file": f})

            # ---------------- PPT → PDF ----------------
            elif mode == "ppt_to_pdf":
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}ppt-to-pdf", files={"file": f})

            # ---------------- Merge PDFs ----------------
            elif mode == "merge_pdfs":
                files_payload = [("files", (Path(f).name, open(f, "rb"))) for f in input_files]
                r = requests.post(f"{url_base}merge-pdfs", files=files_payload)

            # ---------------- Split PDF ----------------
            elif mode == "split_pdf":
                # Get the input from the extra field
                pages_text = getattr(self, "split_pdf_input").text()  # e.g., "1-3"

                # Split into start and end pages
                try:
                    start_page, end_page = [int(p.strip()) for p in pages_text.split('-')]
                except ValueError:
                    input_attr.setText("Invalid page range! Use format start-end, e.g., 1-3")
                    return

                with open(input_files[0], "rb") as f:
                    r = requests.post(
                        f"{url_base}split-pdf",
                        files={"file": f},
                        data={"start_page": start_page, "end_page": end_page}
                    )

            # ---------------- Rotate PDF ----------------
            elif mode == "rotate_pdf":
                angle = getattr(self, "rotate_pdf_input").text()
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}rotate-pdf", files={"file": f}, data={"angle": angle})

            # ---------------- Watermark PDF ----------------
            elif mode == "watermark_pdf":
                watermark_text = getattr(self, "watermark_pdf_input").text()
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}watermark-pdf", files={"file": f}, data={"watermark_text": watermark_text})

            # ---------------- Extract PDF Text ----------------
            elif mode == "extract_text_pdf":
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}extract-text-pdf", files={"file": f})

            # ---------------- OCR PDF ----------------
            elif mode == "ocr_pdf":
                with open(input_files[0], "rb") as f:
                    r = requests.post(f"{url_base}ocr-pdf", files={"file": f})

            # ---------------- Handle Response ----------------
            if r.status_code == 200:
                cd = r.headers.get("Content-Disposition", "")
                if "filename=" in cd:
                    filename = cd.split("filename=")[-1].strip('"')
                else:
                    # Fallback based on conversion mode
                    if mode == "pdf_to_word":
                        filename = Path(input_files[0]).stem + ".docx"
                    elif mode == "pdf_image":
                        filename = Path(input_files[0]).stem + "_images.zip"
                    elif mode == "images_to_pdf":
                        filename = Path(input_files[0]).stem + ".pdf"
                    else:
                        filename = Path(input_files[0]).stem + ".pdf"

                out_file = get_unique_filepath(downloads, filename)
                with open(out_file, "wb") as f:
                    f.write(r.content)
                input_attr.setText(f"Saved: {out_file}")

        except Exception as e:
            input_attr.setText(f"Error: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ConverterApp()
    pixmap = QPixmap(resource_path("File_Converter_Logo.ico"))
    window.show()
    sys.exit(app.exec())