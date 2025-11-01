import sys
import os
import datetime
import traceback

from PyQt6.QtCore import QObject, pyqtSignal
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from iogen_main import Ui_MainWindow

import IO_Table_generator as iogen

# --- Класс для перехвата вывода print() и перенаправления в GUI ---
class EmittingStream(QObject):
    textWritten = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._buffer = ""

    def write(self, text):
        # Добавляем текст в буфер и обрабатываем построчно
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            self.textWritten.emit(line + "\n")

    def flush(self):
        if self._buffer:
            self.textWritten.emit(self._buffer)
            self._buffer = ""

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._input_file_path = None
        self._map_file_path = None
        self._default_dir = ""
        self.setupUi(self)
        self.connectSignalsSlots()
        self.statusbar.showMessage("Start application")

        # === Подключаем перехват stdout ===
        self.emitting_stream = EmittingStream()
        self.emitting_stream.textWritten.connect(self.normalOutputWritten)
        sys.stdout = self.emitting_stream
        sys.stderr = self.emitting_stream

    def normalOutputWritten(self, text):
        """Добавляет текст в QTextEdit с переводами строк"""
        cursor = self.textEdit_log.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        cursor.insertText(text)
        self.textEdit_log.setTextCursor(cursor)
        self.textEdit_log.ensureCursorVisible()

    def connectSignalsSlots(self):
        self.pushButton.clicked.connect(self.onInputFileSelect)
        self.pushButton_3.clicked.connect(self.onMapFileSelect)
        self.pushButton_6.clicked.connect(self.onLoadBtn)

    def onLoadBtn(self):
        """Загружает L5X и карту сигналов через IO_Table_generator"""
        try:
            if not self._input_file_path:
                QMessageBox.warning(self, "Error", "No L5X file selected!")
                return

            if not self._input_file_path.lower().endswith(".l5x"):
                QMessageBox.warning(self, "Error", "Only L5X files are supported in GUI mode.")
                return

            self.statusbar.showMessage("Loading L5X project...")
            QApplication.processEvents()

            print(f"🔹 Starting load: {self._input_file_path}")
            print(f"🔹 Map file: {self._map_file_path or '(none)'}")

            # --- вызов основной функции из IO_Table_generator ---
            iogen.read_input_l5x(
                self._input_file_path,
                map_file_name=self._map_file_path,
                debug=True
            )

            print("✅ File successfully processed.")
            self.statusbar.showMessage("Processing complete ✅")

        except Exception as e:
            err_msg = f"Error while processing:\n{e}\n\n{traceback.format_exc()}"
            self.statusbar.showMessage("Error during processing")
            QMessageBox.critical(self, "Load Error", err_msg)
            print(f"❌ Exception: {e}\n{traceback.format_exc()}")

    def onMapFileSelect(self):
        self.statusbar.showMessage("Select map file")
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select map file...",
            self._default_dir,  # Default directory (пустая строка — домашний каталог пользователя)
            "TXT file (*.txt);;All Files (*)",  # Расширенный фильтр
        )

        if filename:
            self._map_file_path = filename
            # self._default_dir = os.path.dirname(filename)
            self.lineEdit_3.setText(os.path.basename(filename))
            self.statusbar.showMessage(f"Map file selected [{filename}]")

            # --- Получаем информацию о файле ---
            file_info = os.stat(filename)
            file_size_kb = file_info.st_size / 1024
            mod_time = datetime.datetime.fromtimestamp(file_info.st_mtime)

            # --- Формируем красивый текст ---
            file_info_text = (
                f"<b>Map file selected:</b><br>{filename}<br><br>"
                f"<b>Size:</b> {file_info.st_size} Bytes<br>"
                f"<b>Modified:</b> {mod_time.strftime('%Y-%m-%d %H:%M:%S')}<br>"
                f"<b>Type:</b> {os.path.splitext(filename)[1].upper()}<br>"
            )

            # --- Обновляем label ---
            self.label_2.setText(file_info_text)
        else:
            self.statusbar.showMessage("Map file not selected")
            self.label_2.setText("No file selected.")


    def onInputFileSelect(self):
        self.statusbar.showMessage("Select input file")
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select project file...",
            self._default_dir,  # Default directory (пустая строка — домашний каталог пользователя)
            "L5X XML file (*.L5X);;CSV file (*.csv);;All Files (*)",  # Расширенный фильтр
        )

        if filename:
            self._input_file_path = filename
            self._default_dir = os.path.dirname(filename)
            self.lineEdit.setText(os.path.basename(filename))
            self.statusbar.showMessage(f"Input file selected [{filename}]")

            # --- Получаем информацию о файле ---
            file_info = os.stat(filename)
            file_size_kb = file_info.st_size / 1024
            mod_time = datetime.datetime.fromtimestamp(file_info.st_mtime)

            # --- Формируем красивый текст ---
            file_info_text = (
                f"<b>Input file selected:</b><br>{filename}<br><br>"
                f"<b>Size:</b> {file_size_kb:.1f} KB<br>"
                f"<b>Modified:</b> {mod_time.strftime('%Y-%m-%d %H:%M:%S')}<br>"
                f"<b>Type:</b> {os.path.splitext(filename)[1].upper()}<br>"
            )

            # --- Обновляем label ---
            self.label.setText(file_info_text)
        else:
            self.statusbar.showMessage("Input file not selected")
            self.label.setText("No file selected.")


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
