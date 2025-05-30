import ctypes
import os
import sys
import subprocess
from pathlib import Path

from win32mica import ApplyMica, MicaTheme, MicaStyle
from win32com.client import Dispatch
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QLabel, QVBoxLayout, QWidget, QDialog,
    QFileDialog, QCheckBox, QPushButton, QHBoxLayout, QLineEdit, QMessageBox
)
from PySide6.QtCore import Qt, QMimeData, QUrl, Signal
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QMouseEvent


class ShortcutNameDialog(QDialog):
    def __init__(self, default_name, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Customize Shortcut")
        self.setWindowIcon(parent.windowIcon())

        layout = QVBoxLayout(self)

        self.name_label = QLabel("Shortcut name:")
        self.name_input = QLineEdit(default_name)
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_input)

        self.admin_checkbox = QCheckBox("Run as administrator")
        layout.addWidget(self.admin_checkbox)

        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("Create")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button = QPushButton("Cancel")
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

    def get_shortcut_name(self):
        return self.name_input.text().strip()

    def run_as_admin(self):
        return self.admin_checkbox.isChecked()


class DropArea(QLabel):
    doubleClicked = Signal()

    def __init__(self, main_window, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setAcceptDrops(True)
        self.setText("Drag and drop an executable file here\nor double-click to browse")
        self.setStyleSheet("""
            DropArea {
                border: 2px dashed #aaaaaa;
                border-radius: 10px;
                padding: 20px;
            }
            DropArea:hover {
                border-color: #666666;
                background-color: rgba(255, 255, 255, 0.05);
            }
        """)

    def mouseDoubleClickEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            self.doubleClicked.emit()
        super().mouseDoubleClickEvent(event)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if len(urls) == 1 and urls[0].toLocalFile().lower().endswith('.exe'):
                event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        file_path = event.mimeData().urls()[0].toLocalFile()
        self.main_window.process_dropped_file(file_path)


class ShortcutCreatorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_windows_appid()
        self.init_ui()
        self.apply_visual_styles()

    def setup_windows_appid(self):
        self.appid = 'com.yourcompany.shortcutcreator.3.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(self.appid)

    def init_ui(self):
        self.setWindowTitle("Start Menu Shortcut Creator")
        self.setGeometry(100, 100, 450, 250)
        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        self.setFixedSize(self.width(), self.height())
        self.setWindowIcon(self.load_icon('icon.ico'))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.drop_area = DropArea(main_window=self, parent=central_widget)
        self.drop_area.doubleClicked.connect(self.browse_for_exe)
        layout.addWidget(self.drop_area)

    def apply_visual_styles(self):
        self.setAttribute(Qt.WA_TranslucentBackground)
        ApplyMica(
            int(self.winId()),
            MicaTheme.AUTO,
            MicaStyle.ALT,
            OnThemeChange=self.apply_theme_style
        )
        self.apply_theme_style(MicaTheme.DARK)

    def apply_theme_style(self, theme):
        if theme == MicaTheme.DARK:
            self.drop_area.setStyleSheet("""
                DropArea {
                    border: 2px dashed #aaaaaa;
                    border-radius: 10px;
                    padding: 20px;
                    color: white;
                }
                DropArea:hover {
                    border-color: #666666;
                    background-color: rgba(255, 255, 255, 0.05);
                }
            """)
        else:
            self.drop_area.setStyleSheet("""
                DropArea {
                    border: 2px dashed #666666;
                    border-radius: 10px;
                    padding: 20px;
                    color: black;
                }
                DropArea:hover {
                    border-color: #333333;
                    background-color: rgba(0, 0, 0, 0.05);
                }
            """)

    def browse_for_exe(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Executable", "", "Executable Files (*.exe);;All Files (*)")
        if file_path:
            self.process_dropped_file(file_path)

    def load_icon(self, icon_name):
        icon_path = self.resource_path(icon_name)
        return QIcon(icon_path) if os.path.exists(icon_path) else QIcon()

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def process_dropped_file(self, exe_path):
        if not os.path.exists(exe_path):
            self.show_error("File not found", "The selected file doesn't exist.")
            return
        self.prompt_shortcut_options(exe_path)

    def prompt_shortcut_options(self, exe_path):
        exe_name = os.path.basename(exe_path)
        default_name = self.clean_default_name(exe_name)

        dialog = ShortcutNameDialog(default_name, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            shortcut_name = dialog.get_shortcut_name()
            run_as_admin = dialog.run_as_admin()

            if shortcut_name:
                self.create_shortcut(exe_path, shortcut_name, run_as_admin)
                self.drop_area.setText(f"Created: {shortcut_name}")

    def clean_default_name(self, filename):
        name = os.path.splitext(filename)[0]
        for sep in ['_', '-', '.']:
            name = name.replace(sep, ' ')
        name = name.rsplit('v', 1)[0].strip()
        return name

    def create_shortcut(self, exe_path, shortcut_name, run_as_admin):
        try:
            programs_folder = self.get_programs_folder()
            shortcut_path = self.get_shortcut_path(programs_folder, shortcut_name)

            if self.check_existing_shortcut(shortcut_path):
                return

            self.create_windows_shortcut(exe_path, shortcut_path, run_as_admin)
        except Exception as e:
            self.show_error("Shortcut Creation Failed", str(e))

    def get_programs_folder(self):
        return os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Windows', 'Start Menu', 'Programs')

    def get_shortcut_path(self, programs_folder, shortcut_name):
        if not shortcut_name.lower().endswith('.lnk'):
            shortcut_name += ".lnk"
        return os.path.join(programs_folder, shortcut_name)

    def check_existing_shortcut(self, shortcut_path):
        if os.path.exists(shortcut_path):
            reply = QMessageBox.question(
                self,
                "Shortcut Exists",
                f"A shortcut named '{os.path.basename(shortcut_path)}' already exists.\nOverwrite?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            return reply == QMessageBox.StandardButton.No
        return False

    def create_windows_shortcut(self, exe_path, shortcut_path, run_as_admin):
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.Targetpath = exe_path
        shortcut.WorkingDirectory = os.path.dirname(exe_path)

        if run_as_admin:
            temp_path = shortcut_path[:-4] + "_admin.lnk"
            shortcut.save(temp_path)

            ps_script = f"""
            $bytes = [System.IO.File]::ReadAllBytes("{temp_path}")
            $bytes[0x15] = $bytes[0x15] -bor 0x20
            [System.IO.File]::WriteAllBytes("{shortcut_path}", $bytes)
            Remove-Item "{temp_path}"
            """
            subprocess.run(["powershell", "-Command", ps_script], check=True)
        else:
            shortcut.save()

    def show_error(self, title, message):
        self.drop_area.setText("Error creating shortcut")
        QMessageBox.critical(self, title, message, QMessageBox.StandardButton.Ok)


def main():
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icon.ico'))

    window = ShortcutCreatorApp()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
