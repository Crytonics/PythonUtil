import sys
import os
import subprocess

def check_and_install_requirements():
    try:
        import pkg_resources
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'setuptools'])
        import pkg_resources

    subprocess.check_call([sys.executable, '-m', 'pip', 'install', '--upgrade', 'pip'])

    requirements_path = 'requirements.txt'
    if os.path.isfile(requirements_path):
        with open(requirements_path, 'r') as file:
            requirements = file.read().splitlines()
        
        installed_packages = {pkg.key for pkg in pkg_resources.working_set}
        missing_packages = [pkg for pkg in requirements if pkg not in installed_packages]

        if missing_packages:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing_packages])

check_and_install_requirements()

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

        self.installTab = QWidget()
        self.tabWidget.addTab(self.installTab, "Install Programs")

        installLayout = QVBoxLayout()
        hLayout = QHBoxLayout()

        installLayout.addLayout(hLayout)

        self.listWidget = QListWidget()
        installLayout.addWidget(self.listWidget)

        self.installTab.setLayout(installLayout)
        self.setLayout(layout)
        self.loadFolders()
    
    def loadFolders(self):
        folder_path = 'Programs'
        folders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
        self.totalPrograms = len(folders)  # Set the total number of programs
        for folder in folders:
            item = QListWidgetItem(folder)
            self.listWidget.addItem(item)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())