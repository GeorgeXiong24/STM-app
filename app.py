from __future__ import annotations

import shutil
import sys
import tempfile
import random
import json
import html
import os
import ssl
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any

from numbers_parser import Document
from openpyxl import load_workbook
from PySide6.QtCore import QObject, QThread, QTimer, Qt, Signal, Slot
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QHeaderView,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


class AnswerJudgeWorker(QObject):
    finished = Signal(object, str)

    def __init__(self, word: str, explanation: str, answer: str) -> None:
        super().__init__()
        self.word = word
        self.explanation = explanation
        self.answer = answer

    @staticmethod
    def _build_ssl_context(base_url: str) -> ssl.SSLContext:
        verify_ssl = os.environ.get("DEEPSEEK_VERIFY_SSL", "true").lower()
        should_verify = verify_ssl not in {"0", "false", "no", "off"}

        if should_verify:
            return ssl.create_default_context()

        return ssl._create_unverified_context()

    @staticmethod
    def _is_certificate_error(error: BaseException) -> bool:
        message = str(error).lower()
        return (
            "certificate verify failed" in message
            or "self-signed certificate" in message
            or "ssl: cert" in message
            or "unable to get local issuer certificate" in message
        )

    def _send_judgment_request(self, request: urllib.request.Request, ssl_context: ssl.SSLContext):
        with urllib.request.urlopen(request, timeout=30, context=ssl_context) as response:
            result = json.loads(response.read().decode("utf-8"))
        content = result["choices"][0]["message"]["content"].strip()
        if content.startswith("```"):
            content = content.split("\n", 1)[1].rsplit("```", 1)[0].strip()
        judgment = json.loads(content)
        correct = judgment["correct"]
        reason = str(judgment.get("reason", ""))
        if not isinstance(correct, bool):
            raise ValueError("The API returned a non-boolean correctness value.")
        return correct, reason

    def run(self) -> None:
        api_key = "sk-23c4c5238b7646f486855c5bed742e04"
        if not api_key:
            self.finished.emit(False, "DEEPSEEK_API_KEY is not set.")
            return

        base_url = os.environ.get("DEEPSEEK_BASE_URL", "https://api.deepseek.com")
        model = os.environ.get("DEEPSEEK_MODEL", "deepseek-chat")
        payload = {
            "model": model,
            "temperature": 0,
            "response_format": {"type": "json_object"},
            "messages": [
                {
                    "role": "system",
                    "content": (
                        "Judge whether a student's Chinese definition matches the reference "
                        "meaning. Accept synonyms, natural paraphrases, and minor typos. "
                        "Reject unrelated or materially incorrect meanings. Return JSON only "
                        "with boolean 'correct' and short Chinese string 'reason'."
                    ),
                },
                {
                    "role": "user",
                    "content": json.dumps(
                        {
                            "word": self.word,
                            "reference_definition": self.explanation,
                            "student_answer": self.answer,
                        },
                        ensure_ascii=False,
                    ),
                },
            ],
        }
        request = urllib.request.Request(
            f"{base_url.rstrip('/')}/chat/completions",
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        ssl_context = self._build_ssl_context(base_url)
        try:
            correct, reason = self._send_judgment_request(request, ssl_context)
        except urllib.error.HTTPError as error:
            try:
                details = error.read().decode("utf-8", errors="replace")
            except Exception:
                details = str(error)
            self.finished.emit(None, f"DeepSeek HTTP {error.code}: {details}")
            return
        except urllib.error.URLError as error:
            if self._is_certificate_error(error):
                try:
                    correct, reason = self._send_judgment_request(request, ssl._create_unverified_context())
                except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, KeyError, IndexError, ValueError, ssl.SSLError) as retry_error:
                    self.finished.emit(None, f"Could not judge the answer: {retry_error}")
                    return
                self.finished.emit(correct, reason)
                return
            self.finished.emit(None, f"Could not judge the answer: {error}")
            return
        except (TimeoutError, KeyError, IndexError, ValueError, ssl.SSLError) as error:
            if self._is_certificate_error(error):
                try:
                    correct, reason = self._send_judgment_request(request, ssl._create_unverified_context())
                except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError, KeyError, IndexError, ValueError, ssl.SSLError) as retry_error:
                    self.finished.emit(None, f"Could not judge the answer: {retry_error}")
                    return
                self.finished.emit(correct, reason)
                return
            self.finished.emit(None, f"Could not judge the answer: {error}")
            return
        self.finished.emit(correct, reason)


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
            #practiceWord { border: 1px solid #e1ddd4; border-radius: 8px; background: #fffdf9; font: 700 34px 'Avenir Next'; }
            #answerBox { border: 1px solid #c9c1b5; border-radius: 6px; background: #fffdf9; padding: 10px; font: 700 20px 'Avenir Next'; }
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
        entries = [
            (word_item.text(), explanation_item.text())
            for row in range(self.words_table.rowCount())
            if (word_item := self.words_table.item(row, 0)) is not None
            and (explanation_item := self.words_table.item(row, 1)) is not None
        ]
        if not entries:
            return

        random.shuffle(entries)
        self.practice_entries = entries
        self.practice_incorrect_entries: list[tuple[str, str]] = []
        self.practice_given_up_entries: list[tuple[str, str]] = []
        self.practice_error_counts = {entry: 0 for entry in entries}
        self.practice_index = 0
        self.practice_attempts = 0
        self.practice_reviewing = False
        self._show_practice_window(self.practice_entries[0])
        self.menuBar().hide()

    def _show_practice_window(self, entry: tuple[str, str]) -> None:
        word, explanation = entry
        practice_widget = QWidget()
        practice_layout = QVBoxLayout(practice_widget)
        practice_layout.setContentsMargins(34, 28, 34, 28)
        practice_layout.setSpacing(20)

        word_label = QLabel("Word")
        word_label.setObjectName("wordsLabel")
        word_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        practice_layout.addWidget(word_label)
        word_display = QLabel(word)
        word_display.setObjectName("practiceWord")
        word_display.setAlignment(Qt.AlignmentFlag.AlignCenter)
        word_display.setFixedSize(560, 150)
        practice_layout.addWidget(word_display, 0, Qt.AlignmentFlag.AlignHCenter)

        answer_box = QLineEdit()
        answer_box.setPlaceholderText("Enter the Chinese definition")
        answer_box.setObjectName("answerBox")
        answer_box.setInputMethodHints(Qt.InputMethodHint.ImhNone)
        answer_box.setFixedSize(460, 90)
        practice_layout.addWidget(answer_box, 0, Qt.AlignmentFlag.AlignHCenter)
        result_label = QLabel("")
        result_label.setObjectName("statusLabel")
        result_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        result_label.setWordWrap(True)
        practice_layout.addWidget(result_label)
        practice_layout.addStretch(1)

        action_row = QHBoxLayout()
        give_up_button = QPushButton("Give up")
        give_up_button.setObjectName("goButton")
        give_up_button.clicked.connect(self._give_up_current_word)
        action_row.addWidget(give_up_button)
        action_row.addStretch()
        ok_button = QPushButton("OK")
        ok_button.setObjectName("goButton")
        ok_button.clicked.connect(
            lambda: self._submit_answer(word, explanation, answer_box, ok_button)
        )
        answer_box.returnPressed.connect(ok_button.click)
        action_row.addWidget(ok_button)
        practice_layout.addLayout(action_row)

        self.setCentralWidget(practice_widget)
        self.practice_widget = practice_widget
        self.practice_word_label = word_label
        self.practice_word_display = word_display
        self.answer_box = answer_box
        self.ok_button = ok_button
        self.give_up_button = give_up_button
        self.result_label = result_label
        answer_box.setFocus(Qt.FocusReason.OtherFocusReason)
        answer_box.activateWindow()

    def _submit_answer(
        self, word: str, explanation: str, answer_box: QLineEdit, ok_button: QPushButton
    ) -> None:
        answer = answer_box.text().strip()
        if not answer:
            self.result_label.setText("Enter an answer first")
            return
        self.practice_attempts += 1
        answer_box.setReadOnly(True)
        ok_button.setText("Judging...")
        ok_button.setEnabled(False)
        self.result_label.setText("Checking your answer...")

        thread = QThread(self)
        worker = AnswerJudgeWorker(word, explanation, answer)
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.finished.connect(self._handle_judgment, Qt.ConnectionType.QueuedConnection)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        self.answer_judge_thread = thread
        self.answer_judge_worker = worker
        self.answer_judge_context = (answer_box, ok_button, thread, worker)
        thread.start()

    @Slot(object, str)
    def _handle_judgment(self, correct: object, reason: str) -> None:
        context = getattr(self, "answer_judge_context", None)
        if context is None:
            return
        answer_box, ok_button, thread, worker = context
        self._show_judgment(
            correct, reason, answer_box, ok_button, thread, worker
        )
        thread.quit()

    def _show_judgment(
        self,
        correct: object,
        reason: str,
        answer_box: QLineEdit,
        ok_button: QPushButton,
        thread: QThread,
        worker: AnswerJudgeWorker,
    ) -> None:
        if correct is None:
            answer_box.setReadOnly(False)
            answer_box.setStyleSheet("border: 2px solid #c85132;")
            ok_button.setText("Try again")
            ok_button.setEnabled(True)
            self.result_label.setText(reason)
            return

        answer_box.setReadOnly(False)
        answer_box.setStyleSheet(
            "border: 2px solid #3a8f62;" if correct else "border: 2px solid #c85132;"
        )
        if correct:
            self.result_label.setText(reason or "Correct")
            self._advance_after_answer(answer_box, ok_button, thread, worker)
        else:
            entry = self.practice_entries[self.practice_index]
            self.practice_error_counts[entry] += 1
            if self.practice_attempts < 2:
                answer_box.setReadOnly(False)
                answer_box.clear()
                answer_box.setStyleSheet("border: 2px solid #c85132;")
                answer_box.setFocus(Qt.FocusReason.OtherFocusReason)
                ok_button.setText("Try again")
                ok_button.setEnabled(True)
                self.result_label.setText("Incorrect. One attempt left.")
                return

            if entry not in self.practice_incorrect_entries:
                self.practice_incorrect_entries.append(entry)
            answer_box.setReadOnly(True)
            ok_button.setText("Next")
            ok_button.setEnabled(True)
            try:
                ok_button.clicked.disconnect()
            except RuntimeError:
                pass
            ok_button.clicked.connect(self._show_next_practice_entry)
            self.result_label.setText(f"Incorrect twice. Correct answer: {entry[1]}")
        thread.finished.connect(lambda: self._release_judge_references(worker))

    def _advance_after_answer(
        self,
        answer_box: QLineEdit,
        ok_button: QPushButton,
        thread: QThread,
        worker: AnswerJudgeWorker,
    ) -> None:
        QTimer.singleShot(
            900,
            self._show_next_practice_entry,
        )

    def _show_next_practice_entry(self) -> None:
        self.practice_index += 1
        if self.practice_index >= len(self.practice_entries):
            if self.practice_incorrect_entries:
                self.practice_entries = self.practice_incorrect_entries
                self.practice_incorrect_entries = []
                self.practice_index = 0
                self.practice_reviewing = True
            else:
                self._finish_practice()
                return

        self.practice_attempts = 0
        self._show_practice_window(self.practice_entries[self.practice_index])

    def _give_up_current_word(self) -> None:
        entry = self.practice_entries[self.practice_index]
        if entry not in self.practice_given_up_entries:
            self.practice_given_up_entries.append(entry)
        self._show_next_practice_entry()

    def _finish_practice(self) -> None:
        self._show_statistics()

    def _show_statistics(self) -> None:
        statistics = "\n".join(
            f"{html.escape(entry[0])}: {count} incorrect"
            for entry, count in self.practice_error_counts.items()
        )
        given_up = "\n".join(
            f'<span style="color: #c85132;">{html.escape(entry[0])}</span>'
            for entry in self.practice_given_up_entries
        )
        report = f"Finished\n\n{statistics}"
        if given_up:
            report += f"\n\nGiven up\n{given_up}"
        self.practice_word_label.hide()
        self.practice_word_display.hide()
        self.answer_box.hide()
        self.ok_button.setText("Finish")
        self.ok_button.setEnabled(True)
        try:
            self.ok_button.clicked.disconnect()
        except RuntimeError:
            pass
        self.ok_button.clicked.connect(self._return_to_start)
        self.result_label.setStyleSheet("font-size: 20px; font-weight: 600;")
        self.result_label.setTextFormat(Qt.TextFormat.RichText)
        self.result_label.setText(report.replace("\n", "<br>"))

    def _return_to_start(self) -> None:
        if self.current_path is not None:
            try:
                self.current_path.unlink(missing_ok=True)
            except OSError:
                pass
        self.current_path = None
        self.words_table.setRowCount(0)
        self.words_label.setText("Words")
        self.file_label.setText("No file uploaded")
        self.storage_label.setText("Uploaded files are removed when you close WR.")
        self.status_label.setText("Ready")
        self._set_go_file_state(False)
        self.setCentralWidget(self.content_widget)
        self.menuBar().show()

    def _release_judge_references(self, worker: AnswerJudgeWorker) -> None:
        if getattr(self, "answer_judge_worker", None) is worker:
            self.answer_judge_worker = None
            self.answer_judge_thread = None
            self.answer_judge_context = None

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
