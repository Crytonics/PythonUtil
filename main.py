import sys
import os
import subprocess
import winreg
import ctypes
import json
import logging

try:
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
                logging.info(f"Starting installation of {self.program_name}")
                print(f"Starting installation of {self.program_name}")
                subprocess.run(self.program_path, shell=True, check=True)
                self.install_finished.emit(self.program_name, True)
                logging.info(f"Successfully installed {self.program_name}")
                print(f"Successfully installed {self.program_name}")
            except subprocess.CalledProcessError:
                self.install_finished.emit(self.program_name, False)
                logging.error(f"Failed to install {self.program_name}")
                print(f"Failed to install {self.program_name}")
            except OSError as e:
                if e.winerror == 740:
                    self.install_finished.emit(self.program_name, False)
                    logging.error(f"Failed to install {self.program_name} due to insufficient permissions")
                    print(f"Failed to install {self.program_name} due to insufficient permissions")
                else:
                    raise

    class UninstallThread(QThread):
        uninstall_finished = pyqtSignal(str, bool)

        def __init__(self, program_name, is_appx_package=False):
            super().__init__()
            self.program_name = program_name
            self.is_appx_package = is_appx_package

        def run(self):
            try:
                logging.info(f"Starting uninstallation of {self.program_name}")
                print(f"Starting uninstallation of {self.program_name}")
                result = subprocess.run([sys.executable, 'scripts/uninstall.py', self.program_name, str(self.is_appx_package).lower()])
                success = (result.returncode == 0)
                self.uninstall_finished.emit(self.program_name, success)
                if success:
                    logging.info(f"Successfully uninstalled {self.program_name}")
                    print(f"Successfully uninstalled {self.program_name}")
                else:
                    logging.error(f"Failed to uninstall {self.program_name}")
                    print(f"Failed to uninstall {self.program_name}")
            except Exception as e:
                self.uninstall_finished.emit(self.program_name, False)
                logging.error(f"Exception occurred while uninstalling {self.program_name}: {e}")
                print(f"Exception occurred while uninstalling {self.program_name}: {e}")

    class App(QWidget):
        def __init__(self):
            super().__init__()
            self.all_installed_successfully = True  # Flag to track installation success
            self.installCounter = 0  # Initialize the counter
            self.totalPrograms = 0  # Initialize the total number of programs
            self.totalProgramsWinget = 0  # Initialize the total number of programs
            self.installedCountWinget = 0  # Initialize the installed count
            self.initUI()  # Call initUI after initializing attributes

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
            self.loadWingetData()

        def is_appx_package_installed(self, app_name):
            ps_command = f"Get-AppxPackage *{app_name}*"
            result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True, shell=True)
            logging.info(f"Checking if appx package {app_name} is installed: {bool(result.stdout.strip())}")
            print(f"Checking if appx package {app_name} is installed: {bool(result.stdout.strip())}")
            return bool(result.stdout.strip())

        def loadUninstallData(self):
            with open('functions/uninstall/uninstall.json', 'r', encoding='utf-8') as file:
                uninstall_data = json.load(file)
            
                for item in uninstall_data:
                    display_name = item.get('name_program', item['name'])
                    list_item = QListWidgetItem(display_name)
                    list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)  # Enable the checkbox for the item
                    list_item.setCheckState(Qt.Unchecked)
                    list_item.setData(Qt.UserRole, item['name'])  # Store the actual program name for uninstallation
                    self.scriptsListWidget.addItem(list_item)
                    if not self.is_program_installed_for_uninstall(item['name']) and not self.is_appx_package_installed(item['name']):
                        list_item.setBackground(Qt.green)
                        list_item.setFlags(list_item.flags() & ~Qt.ItemIsEnabled)
                        list_item.setText(f"        {display_name} (Not Installed)")  # Indent the program name and checkbox

        def uninstallSelected(self):
            selected_items = [self.scriptsListWidget.item(i) for i in range(self.scriptsListWidget.count()) if self.scriptsListWidget.item(i).checkState() == Qt.Checked]
            
            # Reset text for items marked as failed
            for i in range(self.scriptsListWidget.count()):
                item = self.scriptsListWidget.item(i)
                if "(Failed)" in item.text():
                    item.setText(item.text().replace(" (Failed)", ""))
            
            if selected_items:
                for item in selected_items:
                    item.setBackground(Qt.yellow)
                    item.setText(f"{item.text()} (Uninstalling...)")
                
                self.uninstallQueue = selected_items
                self.uninstallNext()
            else:
                QMessageBox.warning(self, 'No Selection', 'Please select programs to uninstall')

        def uninstallNext(self):
            if self.uninstallQueue:
                item = self.uninstallQueue.pop(0)
                program_name = item.data(Qt.UserRole)  # Retrieve the actual program name for uninstallation
                is_appx_package = self.is_appx_package_installed(program_name)
                self.thread = UninstallThread(program_name, is_appx_package)
                self.thread.uninstall_finished.connect(self.onUninstallFinished)
                self.thread.start()
            else:
                QMessageBox.information(self, 'Uninstall', 'All selected programs have been uninstalled')

        def onUninstallFinished(self, program_name, success):
            for i in range(self.scriptsListWidget.count()):
                item = self.scriptsListWidget.item(i)
                if item.data(Qt.UserRole) == program_name:
                    if success:
                        item.setBackground(Qt.green)
                        item.setCheckState(Qt.Unchecked)
                        item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                        item.setText(f"{item.text()} (Uninstalled)")
                    else:
                        item.setBackground(Qt.red)
                        item.setText(f"{item.text()} (Failed)")
                    break
            self.uninstallNext()

        def updateCounterLabel(self):
            self.counterLabel.setText(f'Installed: {self.installCounter}/{self.totalPrograms}')

        def loadFolders(self):
            folder_path = 'Programs'
            categories = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
            self.totalPrograms = 0  # Reset the total number of programs

            for category in categories:
                category_item = QListWidgetItem(category)
                category_item.setFlags(category_item.flags() | Qt.ItemIsUserCheckable)  # Enable the checkbox for the category item
                category_item.setCheckState(Qt.Unchecked)
                category_item.setBackground(Qt.lightGray)
                self.listWidget.addItem(category_item)

                category_path = os.path.join(folder_path, category)
                programs = [f for f in os.listdir(category_path) if os.path.isdir(os.path.join(category_path, f))]
                self.totalPrograms += len(programs)  # Increment the total number of programs

                for program in programs:
                    program_item = QListWidgetItem(f"{program}")  # Indent the program name and checkbox
                    program_item.setText(f"        {program}")
                    program_item.setFlags(program_item.flags() | Qt.ItemIsUserCheckable)  # Enable the checkbox for the program item
                    program_item.setCheckState(Qt.Unchecked)
                    if self.is_program_installed(program):
                        program_item.setBackground(Qt.green)
                        program_item.setFlags(program_item.flags() & ~Qt.ItemIsEnabled)
                        program_item.setText(f"        {program} (Installed)")  # Indent the program name and checkbox
                    self.listWidget.addItem(program_item)

            self.updateCounterLabel()  # Update the counter label after loading folders

        # //OLD WAY OF CHECKING INSTALLED PROGRAMS
        # def is_program_installed(self, program_name):
        #     uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
        #     with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key) as key:
        #         for i in range(0, winreg.QueryInfoKey(key)[0]):
        #             sub_key_name = winreg.EnumKey(key, i)
        #             with winreg.OpenKey(key, sub_key_name) as sub_key:
        #                 try:
        #                     display_name = winreg.QueryValueEx(sub_key, "DisplayName")[0]
        #                     if program_name.lower() in display_name.lower():
        #                         return True
        #                 except FileNotFoundError:
        #                     pass
        #     return False

        def is_program_installed(self, program_name):
            if not hasattr(self, 'installed_programs'):
                try:
                    print("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                    logging.info("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                    result = subprocess.run(["winget", "list"], capture_output=True, text=True, check=True)
                    self.installed_programs = result.stdout.lower()
                except subprocess.CalledProcessError:
                    self.installed_programs = ""
            
            if program_name.lower() in self.installed_programs:
                print(f"***[Checking Installed Programs]*** {program_name} is installed")
                logging.info(f"***[Checking Installed Programs]*** {program_name} is installed")
            else:
                print(f"***[Checking Installed Programs]*** {program_name} is not installed")
                logging.info(f"***[Winget]*** {program_name} is not installed")
            return program_name.lower() in self.installed_programs

        def is_program_installed_for_uninstall(self, program_name):
            uninstall_key = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key) as key:
                for i in range(0, winreg.QueryInfoKey(key)[0]):
                    sub_key_name = winreg.EnumKey(key, i)
                    with winreg.OpenKey(key, sub_key_name) as sub_key:
                        try:
                            display_name = winreg.QueryValueEx(sub_key, "DisplayName")[0]
                            if program_name.lower() in display_name.lower():
                                return True
                        except FileNotFoundError:
                            pass
            return False
        
        def selectAll(self, state):
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:  # Check if the item is not marked as installed and is not a category
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)

        def selectAllUninstall(self, state):
            for i in range(self.scriptsListWidget.count()):
                item = self.scriptsListWidget.item(i)
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)

        def applyPolicies(self):
            reply = QMessageBox.question(self, 'Confirm', 'Are you sure you want to apply policies?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                logging.info("Applying policies")
                print("Applying policies")
                subprocess.run([sys.executable, 'scripts/policies.py', 'apply'])
                QMessageBox.information(self, 'Policies Applied', 'Policies have been successfully applied.')
                logging.info("Policies applied successfully")
                print("Policies applied successfully")

        def revertPolicies(self):
            reply = QMessageBox.question(self, 'Confirm', 'Are you sure you want to revert policies?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                logging.info("Reverting policies")
                print("Reverting policies")
                subprocess.run([sys.executable, 'scripts/policies.py', 'revert'])
                QMessageBox.information(self, 'Policies Reverted', 'Policies have been successfully reverted.')
                logging.info("Policies reverted successfully")
                print("Policies reverted successfully")

        def installSelected(self):
            selected_items = [self.listWidget.item(i) for i in range(self.listWidget.count()) if self.listWidget.item(i).checkState() == Qt.Checked]
            
            # Reset text for items marked as failed
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                if "(Failed)" in item.text():
                    item.setText(item.text().replace(" (Failed)", ""))
            
            # Filter out category items
            selected_items = [item for item in selected_items if item.background() != Qt.lightGray]
            
            if selected_items:
                self.installQueue = selected_items
                self.progressBar.setMaximum(100)
                self.progressBar.setValue(0)
                self.increment = 100 / len(selected_items)  # Calculate increment based on the number of selected items
                self.updateCounterLabel()  # Update the counter label initially
                self.installNext()
            else:
                QMessageBox.warning(self, 'No Selection', 'Please select programs to install')

        def categorySelectAll(self, category_item):
            category_index = self.listWidget.row(category_item)
            for i in range(category_index + 1, self.listWidget.count()):
                item = self.listWidget.item(i)
                if item.background() == Qt.lightGray:  # Stop if another category is encountered
                    break
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:  # Check if the item is not marked as installed
                    item.setCheckState(category_item.checkState())

        def onItemChanged(self, item):
            if item.background() == Qt.lightGray:  # Check if the item is a category
                self.categorySelectAll(item)

        def installNext(self):
            if self.installQueue:
                item = self.installQueue.pop(0)
                program_name = item.text().strip()  # Remove leading spaces
                category_name = self.getCategoryForProgram(program_name)
                program_folder = os.path.join('Programs', category_name, program_name)
                exe_files = [f for f in os.listdir(program_folder) if f.endswith('.exe')]
                msi_files = [f for f in os.listdir(program_folder) if f.endswith('.msi')]
                msix_files = [f for f in os.listdir(program_folder) if f.endswith('.msix')]

                all_files = exe_files + msi_files + msix_files

                if all_files:
                    for file in all_files:
                        program_path = os.path.join(program_folder, file)
                        self.thread = InstallThread(program_path, program_name)
                        self.thread.install_finished.connect(self.onInstallFinished)
                        self.thread.start()
                        self.thread.wait()  # Wait for the current installation to finish before starting the next

                        # Check for and run the corresponding pyautogui script with elevated privileges
                        pyautogui_script = os.path.join('functions', 'automate', f'auto_{program_name}.py')
                        if os.path.isfile(pyautogui_script):
                            subprocess.Popen([sys.executable, pyautogui_script])
                else:
                    logging.warning(f'No .exe, .msi, or .msix file found in {program_name} folder')
                    print(f'No .exe, .msi, or .msix file found in {program_name} folder')
                    item.setBackground(Qt.red)
                    item.setText(f"        {program_name} (Failed)")  # Indent the program name and checkbox
                    self.all_installed_successfully = False  # Set flag to False if any installation fails
                    QMessageBox.warning(self, 'File Not Found', f'No .exe, .msi, or .msix file found in {program_name} folder')
                    self.installNext()  # Continue with the next item
            else:
                if self.all_installed_successfully:
                    logging.info('All selected programs have been installed')
                    print('All selected programs have been installed')
                    QMessageBox.information(self, 'Install', 'All selected programs have been installed')
                else:
                    logging.warning('Some programs failed to install')
                    print('Some programs failed to install')
                    self.all_installed_successfully = False  # Set flag to False if any installation fails
                    QMessageBox.warning(self, 'Install', 'Some programs failed to install')
                self.progressBar.setValue(100)  # Ensure progress bar is set to 100% when done

        def getCategoryForProgram(self, program_name):
            folder_path = 'Programs'
            categories = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
            for category in categories:
                category_path = os.path.join(folder_path, category)
                programs = [f for f in os.listdir(category_path) if os.path.isdir(os.path.join(category_path, f))]
                if program_name in programs:
                    return category
            return None

        def onInstallFinished(self, program_name, success):
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                item_text = item.text().strip()  # Remove leading and trailing spaces
                if item_text.startswith(program_name):
                    if success:
                        item.setBackground(Qt.green)
                        item.setCheckState(Qt.Unchecked)
                        item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                        item.setText(f"        {program_name} (Installed)")  # Indent the program name and checkbox
                        self.installCounter += 1  # Increment the counter by 1
                    else:
                        item.setBackground(Qt.red)
                        item.setText(f"        {program_name} (Failed)")  # Indent the program name and checkbox
                        self.all_installed_successfully = False  # Set flag to False if any installation fails
                    break
            self.progressBar.setValue(self.progressBar.value() + int(self.increment))
            self.updateCounterLabel()  # Update the counter label after each installation
            self.installNext()  # Continue with the next item

    
        def loadWingetData(self):
            with open('functions/install/winget.json', 'r', encoding='utf-8') as file:
                self.winget_data = json.load(file)
                self.totalProgramsWinget = len(self.winget_data)
                self.updateCounterLabelWinget()
                
                categories = {}
                for program, details in self.winget_data.items():
                    category = details['category']
                    if category not in categories:
                        categories[category] = []
                    categories[category].append(details)
                
                for category, programs in categories.items():
                    category_item = QListWidgetItem(category)
                    category_item.setFlags(category_item.flags() & ~Qt.ItemIsUserCheckable)  # Disable the checkbox for the category item
                    category_item.setBackground(Qt.lightGray)
                    self.listWidgetWinget.addItem(category_item)
                    
                    for details in programs:
                        list_item = QListWidgetItem(details['Name'])
                        list_item.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)  # Enable the checkbox for the item
                        list_item.setCheckState(Qt.Unchecked)
                        list_item.setData(Qt.UserRole, details['Name'])  # Store the actual program name for installation
                        if self.is_program_installed_winget(details['Name']):
                            list_item.setBackground(Qt.green)
                            list_item.setFlags(list_item.flags() & ~Qt.ItemIsEnabled)
                            list_item.setText(f"{details['Name']} (Installed)")
                            self.installedCountWinget += 1
                            self.updateCounterLabelWinget()
                        self.listWidgetWinget.addItem(list_item)
        
        def is_program_installed_winget(self, program_name):
            if not hasattr(self, 'installed_programs'):
                try:
                    print("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                    logging.info("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                    result = subprocess.run(["winget", "list"], capture_output=True, text=True, check=True)
                    self.installed_programs = result.stdout.lower()
                except subprocess.CalledProcessError:
                    self.installed_programs = ""
            
            if program_name.lower() in self.installed_programs:
                print(f"***[Checking Installed Programs]*** {program_name} is installed")
                logging.info(f"***[Checking Installed Programs]*** {program_name} is installed")
            else:
                print(f"***[Checking Installed Programs]*** {program_name} is not installed")
                logging.info(f"***[Checking Installed Programs]*** {program_name} is not installed")
            return program_name.lower() in self.installed_programs

        def selectAllWinget(self, state):
            for i in range(self.listWidgetWinget.count()):
                item = self.listWidgetWinget.item(i)
                if item.background() == Qt.lightGray:  # Check if the item is a category
                    self.categorySelectAllWinget(item, state)
                elif item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)

        def categorySelectAllWinget(self, category_item, state):
            category_index = self.listWidgetWinget.row(category_item)
            for i in range(category_index + 1, self.listWidgetWinget.count()):
                item = self.listWidgetWinget.item(i)
                if item.background() == Qt.lightGray:  # Stop if another category is encountered
                    break
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)

        def onItemChangedWinget(self, item):
            if item.background() == Qt.lightGray:  # Check if the item is a category
                self.categorySelectAllWinget(item)

        def installSelectedWinget(self):
            selected_items = [self.listWidgetWinget.item(i) for i in range(self.listWidgetWinget.count()) if self.listWidgetWinget.item(i).checkState() == Qt.Checked]
            
            if selected_items:
                self.installQueueWinget = selected_items
                self.progressBarWinget.setMaximum(100)
                self.progressBarWinget.setValue(0)
                self.incrementWinget = 100 / len(selected_items)  # Calculate increment based on the number of selected items
                self.installedCountWinget = 0  # Initialize the installed count
                self.totalSelectedWinget = len(selected_items)  # Total number of selected items
                self.updateCounterLabelWinget()  # Update the counter label initially
                self.installNextWinget()
            else:
                QMessageBox.warning(self, 'No Selection', 'Please select programs to install')

        def installNextWinget(self):
            if self.installQueueWinget:
                item = self.installQueueWinget.pop(0)
                program_name = item.data(Qt.UserRole)  # Retrieve the actual program name for installation
                winget_command = self.winget_data[program_name]['winget']
                ps_command = f"winget install {winget_command} --accept-package-agreements --accept-source-agreements"
                try:
                    logging.info(f"***[Winget]*** Installing {program_name} using winget")
                    print(f"***[Winget]*** Installing {program_name} using winget")
                    subprocess.run(["powershell", "-Command", ps_command], check=True)
                    logging.info(f"***[Winget]*** Successfully installed {program_name}")
                    print(f"***[Winget]*** Successfully installed {program_name}")
                    self.onInstallFinishedWinget(program_name, True)
                except subprocess.CalledProcessError:
                    logging.error(f"***[Winget]*** Failed to install {program_name}")
                    print(f"***[Winget]*** Failed to install {program_name}")
                    self.onInstallFinishedWinget(program_name, False)
            else:
                QMessageBox.information(self, 'Install', 'All selected programs have been installed')
                self.progressBarWinget.setValue(100)  # Ensure progress bar is set to 100% when done

        def onInstallFinishedWinget(self, program_name, success):
            for i in range(self.listWidgetWinget.count()):
                item = self.listWidgetWinget.item(i)
                item_text = item.text().strip()  # Remove leading and trailing spaces
                if item_text == self.winget_data[program_name]['Name']:
                    if success:
                        item.setBackground(Qt.green)
                        item.setCheckState(Qt.Unchecked)
                        item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                        item.setText(f"{item.text()} (Installed)")
                        self.installedCountWinget += 1  # Increment the counter by 1
                    else:
                        item.setBackground(Qt.red)
                        item.setText(f"{item.text()} (Failed)")
                    break
            self.progressBarWinget.setValue(self.progressBarWinget.value() + int(self.incrementWinget))
            self.updateCounterLabelWinget()  # Update the counter label after each installation
            self.installNextWinget()  # Continue with the next item
        
        def updateCounterLabelWinget(self):
            self.counterLabelWinget.setText(f'Installed: {self.installedCountWinget}/{self.totalProgramsWinget}')

    if __name__ == '__main__':
        app = QApplication(sys.argv)
        ex = App()
        ex.show()
        sys.exit(app.exec_())

except Exception as e:
    logging.error("An error occurred: %s", e)
    print("An error occurred: %s", e)
    QMessageBox.critical(None, "Error", f"An error occurred: {e}")
    sys.exit(1)