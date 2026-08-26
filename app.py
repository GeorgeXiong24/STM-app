from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path
from typing import Any

from numbers_parser import Document
from openpyxl import load_workbook
from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QHeaderView,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class SpreadsheetWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.current_path: Path | None = None
        self.temporary_directory = tempfile.TemporaryDirectory(prefix="wr-")
        self.setWindowTitle("WR")
        self.setMinimumSize(900, 760)
        self.resize(1180, 980)
        self._build_ui()
        self._apply_styles()

    def _build_ui(self) -> None:
        self.content_widget = QWidget()
        self.setCentralWidget(self.content_widget)
        root = QVBoxLayout(self.content_widget)
        root.setContentsMargins(34, 28, 34, 28)
        root.setSpacing(20)

        header = QHBoxLayout()
        brand = QLabel("WR")
        brand.setObjectName("brand")
        header.addWidget(brand)
        header.addStretch()
        self.open_button = QPushButton("Open sheet")
        self.open_button.setObjectName("primaryButton")
        self.open_button.clicked.connect(self.choose_file)
        header.addWidget(self.open_button)
        root.addLayout(header)

        intro = QVBoxLayout()
        intro.setSpacing(5)
        title = QLabel("Your workspace for sheets.")
        title.setObjectName("title")
        subtitle = QLabel("Upload a Numbers or Excel file for temporary processing.")
        subtitle.setObjectName("subtitle")
        intro.addWidget(title)
        intro.addWidget(subtitle)
        root.addLayout(intro)

        self.drop_zone = QFrame()
        self.drop_zone.setObjectName("dropZone")
        self.drop_zone.setAcceptDrops(True)
        drop_layout = QVBoxLayout(self.drop_zone)
        drop_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.setSpacing(10)
        upload_mark = QLabel("↑")
        upload_mark.setObjectName("uploadMark")
        upload_mark.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.addWidget(upload_mark)
        drop_title = QLabel("Drop a sheet here")
        drop_title.setObjectName("dropTitle")
        drop_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.addWidget(drop_title)
        drop_hint = QLabel(".numbers and .xlsx files")
        drop_hint.setObjectName("dropHint")
        drop_hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        drop_layout.addWidget(drop_hint)
        self.drop_zone.dragEnterEvent = self._drag_enter
        self.drop_zone.dropEvent = self._drop_file
        root.addWidget(self.drop_zone)

        self.file_label = QLabel("No file uploaded")
        self.file_label.setObjectName("fileLabel")
        root.addWidget(self.file_label)

        self.words_label = QLabel("Words")
        self.words_label.setObjectName("wordsLabel")
        root.addWidget(self.words_label)
        self.words_table = QTableWidget(0, 3)
        self.words_table.setObjectName("wordsTable")
        self.words_table.setHorizontalHeaderLabels(["单词", "解释", ""])
        self.words_table.setAlternatingRowColors(True)
        self.words_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.words_table.horizontalHeader().setStretchLastSection(False)
        self.words_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.words_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.words_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        root.addWidget(self.words_table, 1)

        self.storage_label = QLabel("Uploaded files are removed when you close WR.")
        self.storage_label.setObjectName("storageLabel")
        root.addWidget(self.storage_label)

        self.status_label = QLabel("Ready")
        self.status_label.setObjectName("statusLabel")
        root.addWidget(self.status_label)

        bottom_actions = QHBoxLayout()
        bottom_actions.addStretch()
        self.go_button = QPushButton("Go")
        self.go_button.setObjectName("goButton")
        self.go_button.setProperty("hasFile", False)
        self.go_button.clicked.connect(self._clear_window)
        bottom_actions.addWidget(self.go_button)
        root.addLayout(bottom_actions)

        file_menu = self.menuBar().addMenu("File")
        open_action = QAction("Open…", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.choose_file)
        file_menu.addAction(open_action)

    def _apply_styles(self) -> None:
        self.setStyleSheet("""
            QMainWindow, QWidget { background: #f6f5f1; color: #20211f; }
            QMenuBar { background: #f6f5f1; padding: 5px 12px; }
            QMenuBar::item:selected { background: #e8e5dc; }
            #brand { color: #c85132; font: 700 27px 'Avenir Next'; letter-spacing: 2px; }
            #title { color: #20211f; font: 700 30px 'Avenir Next'; }
            #subtitle { color: #70736d; font: 15px 'Avenir Next'; }
            #primaryButton { background: #c85132; color: white; border: 0; border-radius: 6px; padding: 11px 19px; font: 600 14px 'Avenir Next'; }
            #primaryButton:hover { background: #ad4127; }
            #goButton { background: #20211f; color: white; border: 0; border-radius: 6px; padding: 10px 22px; font: 600 14px 'Avenir Next'; }
            #goButton[hasFile="false"] { background: #aaa9a3; }
            #goButton[hasFile="false"]:hover { background: #989791; }
            #goButton[hasFile="true"] { background: #c85132; }
            #goButton[hasFile="true"]:hover { background: #ad4127; }
            #dropZone { min-height: 145px; border: 1.5px dashed #c9c1b5; border-radius: 8px; background: #fbfaf7; }
            #dropZone:hover { border-color: #c85132; background: #fffdf9; }
            #uploadMark { color: #c85132; font: 700 30px 'Avenir Next'; }
            #dropTitle { color: #353633; font: 600 16px 'Avenir Next'; }
            #dropHint, #statusLabel, #storageLabel { color: #85877f; font: 13px 'Avenir Next'; }
            #fileLabel { color: #3f413d; font: 600 14px 'Avenir Next'; padding: 2px 0; }
            #wordsLabel { color: #353633; font: 600 16px 'Avenir Next'; padding-top: 4px; }
            #wordsTable { border: 1px solid #e1ddd4; border-radius: 6px; background: #fffdf9; font: 14px 'Avenir Next'; }
        """)

    def choose_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Open spreadsheet", "", "Sheets (*.numbers *.xlsx)")
        if path:
            self.open_file(Path(path))

    def _drag_enter(self, event: Any) -> None:
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def _drop_file(self, event: Any) -> None:
        urls = event.mimeData().urls()
        if urls:
            self.open_file(Path(urls[0].toLocalFile()))
            event.acceptProposedAction()

    def open_file(self, path: Path) -> None:
        if path.suffix.lower() not in {".numbers", ".xlsx"}:
            self._show_error("WR accepts .numbers and .xlsx files only.")
            return
        try:
            temporary_path = Path(self.temporary_directory.name) / path.name
            shutil.copy2(path, temporary_path)
        except Exception as error:
            self._show_error(f"Could not upload this file.\n\n{error}")
            return
        self.current_path = temporary_path
        self.file_label.setText(f"Uploaded: {path.name}")
        self.storage_label.setText(f"Temporarily stored as {temporary_path.name}")
        try:
            entries = self._read_words(temporary_path)
        except Exception as error:
            self.words_table.setRowCount(0)
            self.words_label.setText("Words")
            self._set_go_file_state(False)
            self._show_error(f"Could not read the 单词 column.\n\n{error}")
            return

        self.words_table.setRowCount(0)
        for word, explanation in entries:
            row = self.words_table.rowCount()
            self.words_table.insertRow(row)
            self.words_table.setItem(row, 0, QTableWidgetItem(word))
            self.words_table.setItem(row, 1, QTableWidgetItem(explanation))
            delete_button = QPushButton("delete")
            delete_button.clicked.connect(self._delete_word_row)
            self.words_table.setCellWidget(row, 2, delete_button)
        self.words_label.setText(f"Words ({len(entries)})")
        self._set_go_file_state(True)
        self.status_label.setText("Upload ready")

    def _clear_window(self) -> None:
        if not self.go_button.property("hasFile"):
            message_box = QMessageBox(self)
            message_box.setWindowTitle("No file")
            message_box.setText("No file is uploaded")
            message_box.setStandardButtons(QMessageBox.StandardButton.NoButton)
            self.no_file_message = message_box
            message_box.show()
            QTimer.singleShot(1500, message_box.accept)
            return
        self.content_widget.hide()
        self.menuBar().hide()

    def _set_go_file_state(self, has_file: bool) -> None:
        self.go_button.setProperty("hasFile", has_file)
        self.go_button.style().unpolish(self.go_button)
        self.go_button.style().polish(self.go_button)

    def _delete_word_row(self) -> None:
        button = self.sender()
        if button is None:
            return
        for row in range(self.words_table.rowCount()):
            if self.words_table.cellWidget(row, 2) is button:
                self.words_table.removeRow(row)
                self.words_label.setText(f"Words ({self.words_table.rowCount()})")
                return

    def _read_words(self, path: Path) -> list[tuple[str, str]]:
        if path.suffix.lower() == ".xlsx":
            workbook = load_workbook(path, read_only=True, data_only=True)
            try:
                row_groups = [
                    list(worksheet.iter_rows(values_only=True))
                    for worksheet in workbook.worksheets
                ]
            finally:
                workbook.close()

            return self._words_from_groups(row_groups)

        document = Document(str(path))
        row_groups = [
            table.rows(values_only=True)
            for sheet in document.sheets
            for table in sheet.tables
        ]
        return self._words_from_groups(row_groups)

    @classmethod
    def _words_from_groups(cls, row_groups: Any) -> list[tuple[str, str]]:
        entries: list[tuple[str, str]] = []
        found_column = False
        for rows in row_groups:
            try:
                entries.extend(cls._words_from_rows(rows))
                found_column = True
            except ValueError:
                continue
        if not found_column:
            raise ValueError('Columns named "单词" and "解释" were not found.')
        return entries

    @staticmethod
    def _words_from_rows(rows: Any) -> list[tuple[str, str]]:
        rows = iter(rows)
        word_index: int | None = None
        explanation_index: int | None = None
        entries: list[tuple[str, str]] = []
        for row in rows:
            values = list(row)
            if word_index is None or explanation_index is None:
                if "单词" not in values or "解释" not in values:
                    continue
                word_index = values.index("单词")
                explanation_index = values.index("解释")
                continue
            if word_index < len(values) and values[word_index] not in (None, ""):
                explanation = values[explanation_index] if explanation_index < len(values) else ""
                entries.append((str(values[word_index]), str(explanation or "")))
        if word_index is None or explanation_index is None:
            raise ValueError('Columns named "单词" and "解释" were not found.')
        return entries

    def closeEvent(self, event: Any) -> None:
        self.temporary_directory.cleanup()
        super().closeEvent(event)

    def _show_error(self, message: str) -> None:
        self.status_label.setText("Could not open file")
        QMessageBox.critical(self, "Unable to open sheet", message)


def main() -> None:
    application = QApplication(sys.argv)
    application.setApplicationName("WR")
    window = SpreadsheetWindow()
    window.show()
    sys.exit(application.exec())


if __name__ == "__main__":
    main()
