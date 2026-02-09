import sys
import os
import datetime
from pathlib import Path
import traceback

from PyQt6.QtCore import QObject, pyqtSignal, QThread, QSettings, QByteArray, Qt
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QDialog, QVBoxLayout, QTextEdit, \
    QPushButton
from iogen_main import Ui_MainWindow

import IO_Table_generator as iogen

company_name = 'github_com_DamirKh_io_ref'

# --- Поток вывода в GUI ---
class EmittingStream(QObject):
    textWritten = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._buffer = ""

    def write(self, text):
        self._buffer += text
        while "\n" in self._buffer:
            line, self._buffer = self._buffer.split("\n", 1)
            self.textWritten.emit(line + "\n")

    def flush(self):
        if self._buffer:
            self.textWritten.emit(self._buffer)
            self._buffer = ""

# --- Рабочий поток, в котором будет выполняться загурзка L5X ---
class LoaderWorker(QObject):
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, input_file, map_file):
        super().__init__()
        self.input_file = input_file
        self.map_file = map_file

    def run(self):
        try:
            print(f"📂 Loading project: {self.input_file}")
            print(f"🗺 Map file: {self.map_file or 'not provided'}")

            iogen.read_input_l5x(
                self.input_file,
                map_file_name=self.map_file,
                debug=True,
            )

            print("✅ Loading completed successfully.")
            self.finished.emit()

        except Exception as e:
            self.error.emit(str(e))


class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._input_file_path = None
        self._map_file_path = None
        self._out_dir = None
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
        self.pushButton_load.clicked.connect(self.onLoadBtn)
        self.pushButton_preview.clicked.connect(self.preview)
        self.pushButton_selectOutDir.clicked.connect(self.onSelect_OutDir)
        self.pushButton_Save.clicked.connect(self.onSave)
        self.pushButton_drop.clicked.connect(self.onDrop)
        self.pushButton_wipeMap.clicked.connect(self.onWipeMap)
        self.checkBox_useKip.checkStateChanged.connect(self.onUseKip)
        self.pushButton_save_as_module(self.onSaveAsModule)

    def onSaveAsModule(self):
        """Запись данных для Python"""
        # --- Проверка наличия данных ---
        if not hasattr(iogen, "io_config") or not len(iogen.io_config):
            QMessageBox.warning(
                self,
                "Nothing to Save",
                "⚠ Нет данных для сохранения. Сначала загрузите L5X и сформируйте таблицу.",
            )
            self.statusbar.showMessage("Save aborted: no data to save")
            return
        try:
            compl = iogen.write_as_python()
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Select file to save data as Python source...",
            self._default_dir,  # Default directory (пустая строка — домашний каталог пользователя)
            "Python file (*.py);;All Files (*)",  # Расширенный фильтр
        )
        if filename:
            with f=open(filename, "w"):
                pas

        else:
            self.statusbar.showMessage("Python file not selected")




    def onUseKip(self, state: Qt.CheckState):
        iogen.use_kip_tag = state is Qt.CheckState.Checked

    def onWipeMap(self):
        reply = QMessageBox.question(
            self,
            "Подтверждение сброса загруженной карты подстановок",
            "⚠ Загруженная карта подстановок будет сброшена.\nПодстановки выполняться не будут.\nПродолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._map_file_path = None
            self.lineEdit_3.clear()
            self.label_2.setText("No file selected.")
            self.statusbar.showMessage("Map data wiped")
            self.pushButton_3.setEnabled(True)
            self.pushButton_wipeMap.setEnabled(False)
        else:
            self.statusbar.showMessage("Сброс карты подстановок отменён пользователем")


    def onDrop(self):
        """Сбрасывает все загруженные данные и очищает интерфейс"""

        reply = QMessageBox.question(
            self,
            "Подтверждение сброса данных",
            "⚠ Все загруженные данные будут удалены.\nПродолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            # --- Очистка данных в iogen ---
            try:
                if hasattr(iogen, "io_config"):
                    iogen.io_config.clear()
                if hasattr(iogen, "io_description"):
                    iogen.io_description.clear()
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Ошибка очистки данных",
                    f"❌ Не удалось сбросить данные:\n{e}",
                )
                self.statusbar.showMessage("Ошибка очистки данных")
                return

            # --- Очистка UI ---
            # self._input_file_path = None
            # self._map_file_path = None
            # self.lineEdit.clear()
            # self.lineEdit_3.clear()
            # self.lineEdit_Out.clear()
            # self.label.setText("No input file selected.")
            # self.label_2.setText("No map file selected.")
            self.textEdit_log.clear()

            # --- Обновление статуса ---
            self.statusbar.showMessage("✅ Данные сброшены")
            QMessageBox.information(
                self,
                "Сброс завершён",
                "✅ Все данные успешно сброшены.",
            )
        else:
            self.statusbar.showMessage("Сброс отменён пользователем")

    def confirm_overwrite(self, filepath: str) -> bool:
        """Проверяет, существует ли файл, и спрашивает подтверждение на перезапись."""
        path = Path(filepath)
        if not path.exists():
            return True  # файла нет, можно записывать

        reply = QMessageBox.question(
            self,
            "Файл уже существует",
            f"Файл '{path.name}' уже существует в:\n{path.parent}\n\nПерезаписать?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            return True
        else:
            self.statusbar.showMessage("💡 Сохранение отменено пользователем")
            return False

    def onSave(self):
        """Сохранение результата в XLSX-файл с проверками"""
        out_path_str = self.lineEdit_Out.text().strip()

        # --- Проверка наличия данных ---
        if not hasattr(iogen, "io_config") or not len(iogen.io_config):
            QMessageBox.warning(
                self,
                "Nothing to Save",
                "⚠ Нет данных для сохранения. Сначала загрузите L5X и сформируйте таблицу.",
            )
            self.statusbar.showMessage("Save aborted: no data to save")
            return

        # --- Проверка пути ---
        if not out_path_str:
            QMessageBox.warning(
                self,
                "No Output File",
                "⚠ Укажите путь для сохранения файла.",
            )
            self.statusbar.showMessage("Save aborted: no output path specified")
            return

        out_path = Path(out_path_str)
        out_dir = out_path.parent

        # --- Проверка существования папки ---
        if not out_dir.exists():
            try:
                out_dir.mkdir(parents=True, exist_ok=True)
                self.statusbar.showMessage(f"Created directory: {out_dir}")
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Directory Error",
                    f"❌ Не удалось создать папку:\n{out_dir}\n\nОшибка: {e}",
                )
                self.statusbar.showMessage("Save aborted: cannot create directory")
                return

        # --- Проверка прав на запись ---
        if not os.access(out_dir, os.W_OK):
            QMessageBox.critical(
                self,
                "Permission Denied",
                f"❌ Нет прав на запись в каталог:\n{out_dir}",
            )
            self.statusbar.showMessage("Save aborted: permission denied")
            return

        if not self.confirm_overwrite(out_path):
            return  # пользователь отказался

        # --- Попытка записи XLSX ---
        try:
            iogen.write_xlsx(str(out_path), save_base=self.checkBox_saveBase.isChecked(), save_descr_as_cell = self.checkBox_saveDescr.isChecked())
            QMessageBox.information(
                self,
                "Save Successful",
                f"✅ Файл успешно сохранён:\n{out_path}",
            )
            self.statusbar.showMessage(f"Saved: {out_path}")
        except Exception as e:
            QMessageBox.critical(
                self,
                "Save Error",
                f"❌ Ошибка при сохранении файла:\n{e}",
            )
            self.statusbar.showMessage("Save failed")

    def onSelect_OutDir(self):
        """Выбор выходного каталога для сохранения XLSX"""
        self.statusbar.showMessage("Select output directory")

        dirname = QFileDialog.getExistingDirectory(
            self,
            "Select output directory...",
            str(self._default_dir) if hasattr(self, "_default_dir") else ""
        )

        if dirname:
            self._out_dir = Path(dirname)

            if not self._input_file_path:
                self.statusbar.showMessage("⚠ Output directory selected, but input file not chosen yet")
                self.lineEdit_Out.setText(str(self._out_dir))
                return

            # создаём полный путь к файлу XLSX
            input_name = Path(self._input_file_path).stem  # имя без расширения
            out_xlsx = self._out_dir / f"{input_name}.xlsx"

            # показываем пользователю
            self.lineEdit_Out.setText(str(out_xlsx))
            self.statusbar.showMessage(f"✅ Output directory selected [{self._out_dir}]")

        else:
            self.statusbar.showMessage("Output directory not selected")

    def preview(self):
        """Показ предварительного просмотра таблицы с запоминанием размера и позиции"""
        try:
            pv = iogen.write_table(print_to_stdout=False)
        except Exception as e:
            self.statusbar.showMessage("❌ Error while generating preview")
            print(f"❌ Exception: {e}")
            return

        # создаём диалог
        dialog = QDialog(self)
        dialog.setWindowTitle("Preview Table")

        layout = QVBoxLayout(dialog)

        text_edit = QTextEdit(dialog)
        text_edit.setReadOnly(True)
        text_edit.setPlainText(pv)

        # используем моноширинный шрифт, чтобы таблица не "расползалась"
        font = text_edit.font()
        font.setFamily("Courier New")
        font.setPointSize(10)
        text_edit.setFont(font)

        layout.addWidget(text_edit)

        close_button = QPushButton("Close", dialog)
        close_button.clicked.connect(dialog.accept)
        layout.addWidget(close_button)

        dialog.setLayout(layout)

        # --- QSettings для сохранения состояния окна ---
        settings = QSettings(company_name, "IO_Generator")

        # восстанавливаем сохранённое состояние, если оно было
        geometry = settings.value("PreviewDialog/geometry")
        if isinstance(geometry, QByteArray):
            dialog.restoreGeometry(geometry)
        else:
            dialog.resize(800, 600)

        # показываем окно и ждём закрытия
        dialog.exec()

        # сохраняем геометрию окна
        settings.setValue("PreviewDialog/geometry", dialog.saveGeometry())

    # --- Запуск обработки в отдельном потоке ---
    def onLoadBtn(self):
        if hasattr(iogen, "io_config") and len(iogen.io_config):
            reply = QMessageBox.question(
                self,
                "Данные уже загружены",
                "⚠ Если Вы загрузите данные еще раз - это приведет к затиранию уже загруженных данных\n Загрузить?",
            )
            if reply == QMessageBox.StandardButton.No:
                self.statusbar.showMessage("Load aborted by user")
                return
        if not self._input_file_path:
            print("⚠ Input file not selected!")
            return

        self.statusbar.showMessage("Loading started...")
        self.pushButton_preview.setEnabled(False)

        # создаём поток и воркер
        self.thread = QThread()
        self.worker = LoaderWorker(self._input_file_path, self._map_file_path)
        self.worker.moveToThread(self.thread)

        # подключаем сигналы
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.onLoadFinished)
        self.worker.error.connect(self.onLoadError)
        self.worker.finished.connect(self.thread.quit)
        self.worker.error.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.worker.error.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        # стартуем
        self.thread.start()

    def onLoadFinished(self):
        self.statusbar.showMessage("✅ Loading completed successfully.")
        self.pushButton_preview.setEnabled(True)

    def onLoadError(self, message):
        self.statusbar.showMessage("❌ Error during loading.")
        print(f"❌ Exception: {message}")

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

            self.pushButton_wipeMap.setEnabled(True)
            self.pushButton_3.setEnabled(False)
            # --- Обновляем label ---
            self.label_2.setText(file_info_text)
        else:
            self._map_file_path = None
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

            if self._out_dir is not None:
                # создаём полный путь к файлу XLSX
                input_name = Path(self._input_file_path).stem  # имя без расширения
                out_xlsx = self._out_dir / f"{input_name}.xlsx"
            else:
                # создаём полный путь к файлу XLSX
                out_xlsx = Path(self._input_file_path).with_suffix('.xlsx')
            self.lineEdit_Out.setText(str(out_xlsx))

        else:
            self.statusbar.showMessage("Input file not selected")
            self.label.setText("No file selected.")

    def closeEvent(self, event):
        settings = QSettings(company_name, "IO_Generator")
        if self._out_dir:
            settings.setValue("out_dir", str(self._out_dir))

        settings.setValue("MainWindow/geometry", self.saveGeometry())
        super().closeEvent(event)

        settings.sync()  # гарантированная запись на диск
        event.accept()

    def showEvent(self, event):
        settings = QSettings(company_name, "IO_Generator")
        # выходная директория
        out_dir_str = settings.value("out_dir", "")
        if out_dir_str:
            self._out_dir = Path(out_dir_str)
            self.lineEdit_Out.setText(str(self._out_dir))
            self.statusbar.showMessage(f"Restored output directory: {self._out_dir}")

        geometry = settings.value("MainWindow/geometry")
        if isinstance(geometry, QByteArray):
            self.restoreGeometry(geometry)
        super().showEvent(event)


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
