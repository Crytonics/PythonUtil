import sys
import os
import subprocess
import ctypes
import logging
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QListWidget, QPushButton, QMessageBox, QListWidgetItem, QProgressBar, QCheckBox, QLabel, QTabWidget

from scripts.winget_manager import loadWingetData, is_program_installed_winget, selectAllWinget, selectAllWingetUpdaUnins, categorySelectAllWinget, categorySelectAllWingetUpdaUnins, onItemChangedWinget, onItemChangedWingetUpdaUnins, installSelectedWinget, installNextWinget, onInstallFinishedWinget, updateCounterLabelWinget, uninstallSelectedWingetUpdaUnins, uninstallNextWinget, onUninstallFinishedWinget, updateSelectedWingetUpdaUnins, updateNextWinget, onUpdateFinishedWinget
from scripts.install_programs_manager import updateCounterLabel, loadFolders, is_program_installed, is_program_installed_for_uninstall, selectAll, selectAllUninstall, installSelected, categorySelectAll, onItemChanged, installNext, getCategoryForProgram, onInstallFinished
from scripts.policies import applyPolicies, revertPolicies  # Import the methods from policies.py
from scripts.uninstall import loadUninstallData, uninstallSelected, uninstallNext, onUninstallFinished  # Import methods from uninstall.py


# Configure logging
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Re-run the program with admin rights
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
                self.all_installed_successfully = True  # Flag to track installation success
                self.installCounter = 0  # Initialize the counter
                self.totalPrograms = 0  # Initialize the total number of programs
                self.totalProgramsWinget = 0  # Initialize the total number of programs
                self.installedCountWinget = 0  # Initialize the installed count
                self.initUI()  # Call initUI after initializing attributes
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

            # New Tab (Policies)
            self.policies = QWidget()
            self.tabWidget.addTab(self.policies, "Policies")

            PoliciesLayout = QVBoxLayout()
            self.policies.setLayout(PoliciesLayout)

            # Add apply policies button
            self.policiesButton = QPushButton('Apply Policies')
            self.policiesButton.clicked.connect(self.applyPolicies)
            PoliciesLayout.addWidget(self.policiesButton)

            # Add revert policies button
            self.policiesButton = QPushButton('Revert Policies')
            self.policiesButton.clicked.connect(self.revertPolicies)
            PoliciesLayout.addWidget(self.policiesButton)

            # New Tab (Uninstall Programs)
            self.newTab = QWidget()
            self.tabWidget.addTab(self.newTab, "Uninstall Programs")

            newTabLayout = QVBoxLayout()
            self.newTab.setLayout(newTabLayout)

            # Add "Select All" checkbox for uninstalling programs
            self.uninstallSelectAllCheckbox = QCheckBox('Select All')
            self.uninstallSelectAllCheckbox.stateChanged.connect(self.selectAllUninstall)
            newTabLayout.addWidget(self.uninstallSelectAllCheckbox)
            
            # Add new list to the "Uninstall Programs" tab
            self.scriptsListWidget = QListWidget()
            newTabLayout.addWidget(self.scriptsListWidget)

            # Add uninstall button
            self.uninstallButton = QPushButton('Uninstall Selected')
            self.uninstallButton.clicked.connect(self.uninstallSelected)
            newTabLayout.addWidget(self.uninstallButton)

            # Load uninstall data from JSON
            self.loadUninstallData()

            # New Tab (Install Programs)
            self.installTab = QWidget()
            self.tabWidget.addTab(self.installTab, "Install Programs")

            installLayout = QVBoxLayout()
            hLayout = QHBoxLayout()

            installLayout.addLayout(hLayout)

            self.listWidget = QListWidget()
            installLayout.addWidget(self.listWidget)
            
            # Add "Select All" checkbox
            self.selectAllCheckbox = QCheckBox('Select All')
            self.selectAllCheckbox.stateChanged.connect(self.selectAll)
            hLayout.addWidget(self.selectAllCheckbox)

            self.selectAllCheckbox.stateChanged.connect(self.selectAll)
            self.listWidget.itemChanged.connect(self.onItemChanged)

            self.counterLabel = QLabel('Installed: 0/0')
            hLayout.addWidget(self.counterLabel)

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

            # New Tab (IWinget Intall)
            self.wingetTab = QWidget()
            self.tabWidget.addTab(self.wingetTab, "Winget Install")

            WingetLayout = QVBoxLayout()
            WingethLayout = QHBoxLayout()

            WingetLayout.addLayout(WingethLayout)

            self.listWidgetWinget = QListWidget()
            WingetLayout.addWidget(self.listWidgetWinget)
            
            # Add "Select All" checkbox
            self.selectAllCheckboxWinget = QCheckBox('Select All')
            self.selectAllCheckboxWinget.stateChanged.connect(self.selectAllWinget)
            WingethLayout.addWidget(self.selectAllCheckboxWinget)

            self.selectAllCheckboxWinget.stateChanged.connect(self.selectAllWinget)
            self.listWidgetWinget.itemChanged.connect(self.onItemChangedWinget)

            self.counterLabelWinget = QLabel('Installed: 0/0')
            WingethLayout.addWidget(self.counterLabelWinget)

            # Add install button
            self.installButtonWinget = QPushButton('Install Selected')
            self.installButtonWinget.clicked.connect(self.installSelectedWinget)
            WingetLayout.addWidget(self.installButtonWinget)

            # Add progress bar
            self.progressBarWinget = QProgressBar()
            WingetLayout.addWidget(self.progressBarWinget)
            self.progressBarWinget.setValue(0)

            self.wingetTab.setLayout(WingetLayout)
            self.setLayout(layout)
            #self.loadWingetData()

            # New Tab (IWinget Intall)
            self.wingetUpdaUninsTab = QWidget()
            self.tabWidget.addTab(self.wingetUpdaUninsTab, "Winget Update/Uninstall")

            WingetUpdaUninsLayout = QVBoxLayout()
            WingethUpdaUninsLayout = QHBoxLayout()
            WingetbUpdaUninsLayout = QHBoxLayout()

            WingetUpdaUninsLayout.addLayout(WingethUpdaUninsLayout)

            self.listWidgetWingetUpdaUnins = QListWidget()
            WingetUpdaUninsLayout.addWidget(self.listWidgetWingetUpdaUnins)

            WingetUpdaUninsLayout.addLayout(WingetbUpdaUninsLayout)

            # Add "Select All" checkbox
            self.selectAllCheckboxWingetUpdaUnins = QCheckBox('Select All')
            self.selectAllCheckboxWingetUpdaUnins.stateChanged.connect(self.selectAllWingetUpdaUnins)

            WingethUpdaUninsLayout.addWidget(self.selectAllCheckboxWingetUpdaUnins)

            self.selectAllCheckboxWingetUpdaUnins.stateChanged.connect(self.selectAllWingetUpdaUnins)
            self.listWidgetWingetUpdaUnins.itemChanged.connect(self.onItemChangedWingetUpdaUnins)

            # Add install button
            self.updateButtonWingetUpdaUnins = QPushButton('Update Selected')
            self.updateButtonWingetUpdaUnins.clicked.connect(self.updateSelectedWingetUpdaUnins)
            WingetbUpdaUninsLayout.addWidget(self.updateButtonWingetUpdaUnins)

            # Add install button
            self.uninstallButtonWingetUpdaUnins = QPushButton('Uninstall Selected')
            self.uninstallButtonWingetUpdaUnins.clicked.connect(self.uninstallSelectedWingetUpdaUnins)
            WingetbUpdaUninsLayout.addWidget(self.uninstallButtonWingetUpdaUnins)

            # Add progress bar
            self.progressBarWingetUpdaUnins = QProgressBar()
            WingetUpdaUninsLayout.addWidget(self.progressBarWingetUpdaUnins)
            self.progressBarWingetUpdaUnins.setValue(0)

            self.wingetUpdaUninsTab.setLayout(WingetUpdaUninsLayout)
            self.setLayout(layout)
            self.loadWingetData()

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
        
        # Import methods from policies
        applyPolicies = applyPolicies
        revertPolicies = revertPolicies

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