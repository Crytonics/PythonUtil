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
        
        # Add "Select All" checkbox
        self.selectAllCheckbox = QCheckBox('Select All')
        self.selectAllCheckbox.stateChanged.connect(self.selectAll)
        hLayout.addWidget(self.selectAllCheckbox)

        self.counterLabel = QLabel('Installed: 0/0')
        hLayout.addWidget(self.counterLabel)

        self.listWidget = QListWidget()
        installLayout.addWidget(self.listWidget)

        # Add install button
        self.installButton = QPushButton('Install Selected')
        self.installButton.clicked.connect(self.installSelected)
        installLayout.addWidget(self.installButton)

        # Add progress bar
        self.progressBar = QProgressBar()
        installLayout.addWidget(self.progressBar)
        self.progressBar.setValue(0)

        self.installTab.setLayout(installLayout)
        self.setLayout(layout)
        self.loadFolders()

        # New Tab
        self.newTab = QWidget()
        self.tabWidget.addTab(self.newTab, "Scripts")

        newTabLayout = QVBoxLayout()
        self.newTab.setLayout(newTabLayout)
        
        # Add new list to the "Scripts" tab
        self.scriptsListWidget = QListWidget()
        newTabLayout.addWidget(self.scriptsListWidget)

    def updateCounterLabel(self):
        self.counterLabel.setText(f'Installed: {self.installCounter}/{self.totalPrograms}')

    def loadFolders(self):
        folder_path = 'Programs'
        folders = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
        self.totalPrograms = len(folders)  # Set the total number of programs
        for folder in folders:
            item = QListWidgetItem(folder)
            item.setCheckState(Qt.Unchecked)
            self.listWidget.addItem(item)
        self.updateCounterLabel()  # Update the counter label after loading folders

    def selectAll(self, state):
        for i in range(self.listWidget.count()):
            item = self.listWidget.item(i)
            item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)

    def installSelected(self):
        selected_items = [self.listWidget.item(i) for i in range(self.listWidget.count()) if self.listWidget.item(i).checkState() == Qt.Checked]
        if selected_items:
            self.installQueue = selected_items
            self.progressBar.setMaximum(100)
            self.progressBar.setValue(0)
            self.increment = 100 / len(selected_items)  # Calculate increment based on the number of selected items
            self.updateCounterLabel()  # Update the counter label initially
            self.installNext()
        else:
            QMessageBox.warning(self, 'No Selection', 'Please select programs to install')

    def installNext(self):
        if self.installQueue:
            item = self.installQueue.pop(0)
            program_name = item.text()
            program_folder = os.path.join('Programs', program_name)
            exe_files = [f for f in os.listdir(program_folder) if f.endswith('.exe')]

            if exe_files:
                program_path = os.path.join(program_folder, exe_files[0])

                self.thread = InstallThread(program_path, program_name)
                self.thread.install_finished.connect(self.onInstallFinished)
                self.thread.start()
            else:
                QMessageBox.warning(self, 'File Not Found', f'No .exe file found in {program_name} folder')
                self.installNext()  # Continue with the next item
        else:
            if self.all_installed_successfully:
                QMessageBox.information(self, 'Install', 'All selected programs have been installed')
            else:
                QMessageBox.warning(self, 'Install', 'Some programs failed to install')

    def onInstallFinished(self, program_name, success):
        for i in range(self.listWidget.count()):
            item = self.listWidget.item(i)
            if item.text() == program_name:
                if success:
                    QMessageBox.information(self, 'Install', f'Installing {program_name} completed successfully')
                    item.setBackground(Qt.green)
                    # Disable the checkbox
                    item.setCheckState(Qt.Unchecked)
                    item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                    self.installCounter += 1  # Increment the counter by 1
                else:
                    item.setBackground(Qt.red)
                    QMessageBox.warning(self, 'Install', f'Installing {program_name} failed')
                    self.all_installed_successfully = False  # Set flag to False if any installation fails
                break
        self.progressBar.setValue(self.progressBar.value() + int(self.increment))
        self.updateCounterLabel()  # Update the counter label after each installation
        self.installNext()  # Continue with the next item

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())