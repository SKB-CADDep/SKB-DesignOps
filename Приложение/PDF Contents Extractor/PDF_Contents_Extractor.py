import sys

from pypdf import PdfReader
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_STYLESHEET = """
    QWidget#centralWidget {
        background-color: #f5f6f8;
    }
    QFrame#card {
        background-color: #ffffff;
        border: 1px solid #d7dde5;
        border-radius: 10px;
        padding: 10px;
    }
    QLabel#sectionTitle {
        font-weight: 700;
        font-size: 14px;
        color: #1f2937;
    }
    QLineEdit {
        border: 1px solid #d7dde5;
        border-radius: 6px;
        padding: 6px 10px;
        background-color: #ffffff;
        font-size: 13px;
    }
    QLineEdit:read-only {
        background-color: #f9fafb;
        color: #374151;
    }
    QTextEdit {
        border: 1px solid #d7dde5;
        border-radius: 6px;
        padding: 8px;
        background-color: #ffffff;
        font-size: 13px;
        color: #1f2937;
    }
    QPushButton#loadBtn {
        background-color: #0d6efd;
        color: white;
        font-weight: 600;
        border: none;
        border-radius: 6px;
        padding: 8px 16px;
        min-width: 140px;
    }
    QPushButton#loadBtn:hover {
        background-color: #0b5ed7;
    }
"""


def extract_bookmarks_from_pdf(file_path):
    reader = PdfReader(file_path)
    outlines = reader.outline
    bookmarks = []
    _extract_recursive(outlines, bookmarks)
    return "\n".join(bookmarks)


def _extract_recursive(outlines, bookmarks_list):
    for item in outlines:
        if isinstance(item, list):
            _extract_recursive(item, bookmarks_list)
        else:
            try:
                title = item.title.strip()
                if title:
                    bookmarks_list.append(title)
            except AttributeError:
                pass


class PDFContentsExtractor(QMainWindow):
    def __init__(self, initial_path=None):
        super().__init__()
        self.setWindowTitle("Извлечение содержания PDF")
        self.resize(800, 600)

        self.file_path = None

        central = QWidget()
        central.setObjectName("centralWidget")
        self.setCentralWidget(central)

        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        file_card = QFrame()
        file_card.setObjectName("card")
        file_layout = QVBoxLayout(file_card)

        file_title = QLabel("PDF-файл")
        file_title.setObjectName("sectionTitle")
        file_layout.addWidget(file_title)

        file_row = QHBoxLayout()
        self.file_entry = QLineEdit()
        self.file_entry.setReadOnly(True)
        self.file_entry.setPlaceholderText("Выберите PDF-файл...")
        file_row.addWidget(self.file_entry, stretch=1)

        self.load_button = QPushButton("📂 Загрузить PDF")
        self.load_button.setObjectName("loadBtn")
        self.load_button.setCursor(Qt.PointingHandCursor)
        self.load_button.clicked.connect(self.load_pdf)
        file_row.addWidget(self.load_button)

        file_layout.addLayout(file_row)
        layout.addWidget(file_card)

        content_card = QFrame()
        content_card.setObjectName("card")
        content_layout = QVBoxLayout(content_card)

        content_title = QLabel("Содержание")
        content_title.setObjectName("sectionTitle")
        content_layout.addWidget(content_title)

        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("Содержание появится здесь после загрузки PDF...")
        content_layout.addWidget(self.text_area)

        layout.addWidget(content_card, stretch=1)

        self.setStyleSheet(APP_STYLESHEET)

        if initial_path:
            self.set_pdf_file(initial_path)

    def load_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите PDF-файл",
            "",
            "PDF files (*.pdf)",
        )
        if file_path:
            self.set_pdf_file(file_path)

    def set_pdf_file(self, file_path):
        self.file_path = file_path
        self.file_entry.setText(file_path)
        self.extract_bookmarks()

    def extract_bookmarks(self):
        if not self.file_path:
            return

        try:
            result = extract_bookmarks_from_pdf(self.file_path)
            self.text_area.setPlainText(result)
            if not result.strip():
                self.text_area.setPlainText("В этом PDF не найдено закладок (содержания).")
        except Exception as exc:
            self.text_area.clear()
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Ошибка при обработке PDF:\n{exc}",
            )


def main():
    initial_path = sys.argv[1] if len(sys.argv) > 1 else None

    app = QApplication(sys.argv)
    app.setStyle("windowsvista")

    window = PDFContentsExtractor(initial_path=initial_path)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
