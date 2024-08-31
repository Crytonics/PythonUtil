import sys
import os
import subprocess
import ctypes
import logging
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QListWidget, QPushButton, QMessageBox, QProgressBar, QCheckBox, QLabel, QTabWidget

from scripts.winget_manager import (
    loadWingetData, is_program_installed_winget, selectAllWinget, selectAllWingetUpdaUnins, 
    categorySelectAllWinget, categorySelectAllWingetUpdaUnins, onItemChangedWinget, 
    onItemChangedWingetUpdaUnins, installSelectedWinget, installNextWinget, 
    onInstallFinishedWinget, updateCounterLabelWinget, uninstallSelectedWingetUpdaUnins, 
    uninstallNextWinget, onUninstallFinishedWinget, updateSelectedWingetUpdaUnins, 
    updateNextWinget, onUpdateFinishedWinget
)
from scripts.install_programs_manager import (
    updateCounterLabel, loadFolders, is_program_installed, is_program_installed_for_uninstall, 
    selectAll, selectAllUninstall, installSelected, categorySelectAll, onItemChanged, 
    installNext, getCategoryForProgram, onInstallFinished
)
from scripts.policies import applyPolicies, revertPolicies, installPythonModules
from scripts.uninstall import loadUninstallData, uninstallSelected, uninstallNext, onUninstallFinished

# Configure logging
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

def check_and_install_requirements():
    logging.info("\n\nAPPLICATION STARTED")
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
    print("Winget update list")
    logging.info("Winget update list")
    subprocess.run(["powershell", "-Command", "winget update"])
    print("Winget update completed")
    logging.info("Winget update completed")

check_and_install_requirements()

class App(QWidget):
    def __init__(self):
        try:
            super().__init__()
            self.all_installed_successfully = True
            self.installCounter = 0
            self.totalPrograms = 0
            self.totalProgramsWinget = 0
            self.installedCountWinget = 0
            self.initUI()
        except Exception as e:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
            line_number = exc_tb.tb_lineno
            error_message = f"An error occurred in {fname} at line {line_number}: {e}"
            logging.error(error_message)
            print(error_message)
            QMessageBox.critical(None, "Error", error_message)
            sys.exit(1)

    def initUI(self):
        self.setWindowTitle('Program Manager')
        self.setGeometry(100, 100, 400, 300)

        layout = QVBoxLayout()
        self.tabWidget = QTabWidget()
        layout.addWidget(self.tabWidget)

        self.initScriptsTab()
        self.initUninstallTab()
        self.initInstallTab()
        self.initWingetInstallTab()
        self.initWingetUpdateUninstallTab()

        self.setLayout(layout)
        self.loadFolders()
        self.loadWingetData()

    def initScriptsTab(self):
        self.scripts = QWidget()
        self.tabWidget.addTab(self.scripts, "Scripts")
        scriptsLayout = QVBoxLayout()
        self.scripts.setLayout(scriptsLayout)

        #self.addButton(scriptsLayout, 'Apply group policies', self.applyPolicies)
        #self.addButton(scriptsLayout, 'Revert group policies', self.revertPolicies)
        self.addButton(scriptsLayout, 'Install python modules', self.installPythonModules)

    def initUninstallTab(self):
        self.uninstallTab = QWidget()
        self.tabWidget.addTab(self.uninstallTab, "Uninstall Programs")
        uninstallLayout = QVBoxLayout()
        self.uninstallTab.setLayout(uninstallLayout)

        self.uninstallSelectAllCheckbox = QCheckBox('Select All')
        self.uninstallSelectAllCheckbox.stateChanged.connect(self.selectAllUninstall)
        uninstallLayout.addWidget(self.uninstallSelectAllCheckbox)

        self.scriptsListWidget = QListWidget()
        uninstallLayout.addWidget(self.scriptsListWidget)

        self.uninstallButton = QPushButton('Uninstall Selected')
        self.uninstallButton.clicked.connect(self.uninstallSelected)
        uninstallLayout.addWidget(self.uninstallButton)

        self.loadUninstallData()

    def initInstallTab(self):
        self.installTab = QWidget()
        self.tabWidget.addTab(self.installTab, "Install Programs")
        installLayout = QVBoxLayout()
        hLayout = QHBoxLayout()
        installLayout.addLayout(hLayout)

        self.listWidget = QListWidget()
        installLayout.addWidget(self.listWidget)

        self.selectAllCheckbox = QCheckBox('Select All')
        self.selectAllCheckbox.stateChanged.connect(self.selectAll)
        hLayout.addWidget(self.selectAllCheckbox)

        self.listWidget.itemChanged.connect(self.onItemChanged)

        self.counterLabel = QLabel('Installed: 0/0')
        hLayout.addWidget(self.counterLabel)

        self.installButton = QPushButton('Install Selected')
        self.installButton.clicked.connect(self.installSelected)
        installLayout.addWidget(self.installButton)

        self.progressBar = QProgressBar()
        installLayout.addWidget(self.progressBar)
        self.progressBar.setValue(0)

        self.installTab.setLayout(installLayout)

    def initWingetInstallTab(self):
        self.wingetTab = QWidget()
        self.tabWidget.addTab(self.wingetTab, "Winget Install")
        wingetLayout = QVBoxLayout()
        wingetHLayout = QHBoxLayout()
        wingetLayout.addLayout(wingetHLayout)

        self.listWidgetWinget = QListWidget()
        wingetLayout.addWidget(self.listWidgetWinget)

        self.selectAllCheckboxWinget = QCheckBox('Select All')
        self.selectAllCheckboxWinget.stateChanged.connect(self.selectAllWinget)
        wingetHLayout.addWidget(self.selectAllCheckboxWinget)

        self.listWidgetWinget.itemChanged.connect(self.onItemChangedWinget)

        self.counterLabelWinget = QLabel('Installed: 0/0')
        wingetHLayout.addWidget(self.counterLabelWinget)

        self.installButtonWinget = QPushButton('Install Selected')
        self.installButtonWinget.clicked.connect(self.installSelectedWinget)
        wingetLayout.addWidget(self.installButtonWinget)

        self.progressBarWinget = QProgressBar()
        wingetLayout.addWidget(self.progressBarWinget)
        self.progressBarWinget.setValue(0)

        self.wingetTab.setLayout(wingetLayout)

    def initWingetUpdateUninstallTab(self):
        self.wingetUpdaUninsTab = QWidget()
        self.tabWidget.addTab(self.wingetUpdaUninsTab, "Winget Update/Uninstall")
        wingetUpdaUninsLayout = QVBoxLayout()
        wingetHUpdaUninsLayout = QHBoxLayout()
        wingetBUpdaUninsLayout = QHBoxLayout()
        wingetUpdaUninsLayout.addLayout(wingetHUpdaUninsLayout)

        self.listWidgetWingetUpdaUnins = QListWidget()
        wingetUpdaUninsLayout.addWidget(self.listWidgetWingetUpdaUnins)
        wingetUpdaUninsLayout.addLayout(wingetBUpdaUninsLayout)

        self.selectAllCheckboxWingetUpdaUnins = QCheckBox('Select All')
        self.selectAllCheckboxWingetUpdaUnins.stateChanged.connect(self.selectAllWingetUpdaUnins)
        wingetHUpdaUninsLayout.addWidget(self.selectAllCheckboxWingetUpdaUnins)

        self.listWidgetWingetUpdaUnins.itemChanged.connect(self.onItemChangedWingetUpdaUnins)

        self.updateButtonWingetUpdaUnins = QPushButton('Update Selected')
        self.updateButtonWingetUpdaUnins.clicked.connect(self.updateSelectedWingetUpdaUnins)
        wingetBUpdaUninsLayout.addWidget(self.updateButtonWingetUpdaUnins)

        self.uninstallButtonWingetUpdaUnins = QPushButton('Uninstall Selected')
        self.uninstallButtonWingetUpdaUnins.clicked.connect(self.uninstallSelectedWingetUpdaUnins)
        wingetBUpdaUninsLayout.addWidget(self.uninstallButtonWingetUpdaUnins)

        self.progressBarWingetUpdaUnins = QProgressBar()
        wingetUpdaUninsLayout.addWidget(self.progressBarWingetUpdaUnins)
        self.progressBarWingetUpdaUnins.setValue(0)

        self.wingetUpdaUninsTab.setLayout(wingetUpdaUninsLayout)

    def addButton(self, layout, text, callback):
        button = QPushButton(text)
        button.clicked.connect(callback)
        layout.addWidget(button)

    # Import methods from winget_manager (Update/Uninstall)
    selectAllWingetUpdaUnins = selectAllWingetUpdaUnins
    categorySelectAllWingetUpdaUnins = categorySelectAllWingetUpdaUnins
    onItemChangedWingetUpdaUnins = onItemChangedWingetUpdaUnins
    uninstallSelectedWingetUpdaUnins = uninstallSelectedWingetUpdaUnins
    uninstallNextWinget = uninstallNextWinget
    onUninstallFinishedWinget = onUninstallFinishedWinget
    updateSelectedWingetUpdaUnins = updateSelectedWingetUpdaUnins
    updateNextWinget = updateNextWinget
    onUpdateFinishedWinget = onUpdateFinishedWinget

    # Import methods from uninstall
    loadUninstallData = loadUninstallData
    uninstallSelected = uninstallSelected
    uninstallNext = uninstallNext
    onUninstallFinished = onUninstallFinished

    # Import methods from scripts
    applyPolicies = applyPolicies
    revertPolicies = revertPolicies
    installPythonModules = installPythonModules

    # Import methods from program_manager
    updateCounterLabel = updateCounterLabel
    loadFolders = loadFolders
    is_program_installed = is_program_installed
    is_program_installed_for_uninstall = is_program_installed_for_uninstall
    selectAll = selectAll
    selectAllUninstall = selectAllUninstall
    installSelected = installSelected
    categorySelectAll = categorySelectAll
    onItemChanged = onItemChanged
    installNext = installNext
    getCategoryForProgram = getCategoryForProgram
    onInstallFinished = onInstallFinished

    # Import methods from winget_manager
    loadWingetData = loadWingetData
    is_program_installed_winget = is_program_installed_winget
    selectAllWinget = selectAllWinget
    categorySelectAllWinget = categorySelectAllWinget
    onItemChangedWinget = onItemChangedWinget
    installSelectedWinget = installSelectedWinget
    installNextWinget = installNextWinget
    onInstallFinishedWinget = onInstallFinishedWinget
    updateCounterLabelWinget = updateCounterLabelWinget

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())