import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QListWidget, QPushButton, QMessageBox, QListWidgetItem, QProgressBar, QCheckBox, QLabel, QTabWidget
from PyQt5.QtCore import QThread, pyqtSignal, Qt

class InstallThread(QThread):
    install_finished = pyqtSignal(str, bool)

    def __init__(self, program_path, program_name):
        super().__init__()
        self.program_path = program_path
        self.program_name = program_name
        
    def run(self):
        import subprocess
        try:
            subprocess.run(self.program_path, shell=True, check=True)
            self.install_finished.emit(self.program_name, True)
        except subprocess.CalledProcessError:
            self.install_finished.emit(self.program_name, False)
        except OSError as e:
            if e.winerror == 740:
                self.install_finished.emit(self.program_name, False)
            else:
                raise

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.all_installed_successfully = True  # Flag to track installation success
        self.installCounter = 0  # Initialize the counter
        self.totalPrograms = 0  # Initialize the total number of programs
        self.initUI()  # Call initUI after initializing attributes

    def initUI(self):
        self.setWindowTitle('Program Manager')
        self.setGeometry(100, 100, 400, 300)

        layout = QVBoxLayout()
        self.tabWidget = QTabWidget()
        layout.addWidget(self.tabWidget)

        self.setLayout(layout)
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())