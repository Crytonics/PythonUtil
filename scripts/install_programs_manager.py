import os
import subprocess
import logging
import winreg
import sys
import os
from PyQt5.QtWidgets import QListWidgetItem, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# Configure logging
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def handle_exception(e):
    exc_type, exc_obj, exc_tb = sys.exc_info()
    fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
    line_number = exc_tb.tb_lineno
    error_message = f"An error occurred in {fname} at line {line_number}: {e}"
    
    logging.error(error_message)
    print(error_message)
    QMessageBox.critical(None, "Error", error_message)

try:
    # Define InstallThread class
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

    def updateCounterLabel(self):
        try:
            self.counterLabel.setText(f'Installed: {self.installCounter}/{self.totalPrograms}')
        except Exception as e:
            handle_exception(e)

    def loadFolders(self):
        try:
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
        except Exception as e:
            handle_exception(e)

    def is_program_installed(self, program_name):
        try:
            if not hasattr(self, 'installed_programs'):
                print("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                logging.info("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                result = subprocess.run(["winget", "list"], capture_output=True, text=True, check=True)
                self.installed_programs = result.stdout.lower()
            if program_name.lower() in self.installed_programs:
                print(f"***[Checking Installed Programs]*** {program_name} is installed")
                logging.info(f"***[Checking Installed Programs]*** {program_name} is installed")
                self.installCounter += 1  # Increment the counter by 1
            else:
                print(f"***[Checking Installed Programs]*** {program_name} is not installed")
                logging.info(f"{program_name} is not installed")
            return program_name.lower() in self.installed_programs
        except Exception as e:
            handle_exception(e)

    def is_program_installed_for_uninstall(self, program_name):
        try:
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
        except Exception as e:
            handle_exception(e)

    def selectAll(self, state):
        try:
            for i in range(self.listWidget.count()):
                item = self.listWidget.item(i)
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:  # Check if the item is not marked as installed and is not a category
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)
        except Exception as e:
            handle_exception(e)

    def selectAllUninstall(self, state):
        try:
            for i in range(self.scriptsListWidget.count()):
                item = self.scriptsListWidget.item(i)
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)
        except Exception as e:
            handle_exception(e)

    def installSelected(self):
        try:
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
        except Exception as e:
            handle_exception(e)

    def categorySelectAll(self, category_item):
        try:
            category_index = self.listWidget.row(category_item)
            for i in range(category_index + 1, self.listWidget.count()):
                item = self.listWidget.item(i)
                if item.background() == Qt.lightGray:  # Stop if another category is encountered
                    break
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:  # Check if the item is not marked as installed
                    item.setCheckState(category_item.checkState())
        except Exception as e:
            handle_exception(e)

    def onItemChanged(self, item):
        try:
            if item.background() == Qt.lightGray:  # Check if the item is a category
                self.categorySelectAll(item)
        except Exception as e:
            handle_exception(e)

    def installNext(self):
        try:
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
        except Exception as e:
            handle_exception(e)

    def getCategoryForProgram(self, program_name):
        try:
            folder_path = 'Programs'
            categories = [f for f in os.listdir(folder_path) if os.path.isdir(os.path.join(folder_path, f))]
            for category in categories:
                category_path = os.path.join(folder_path, category)
                programs = [f for f in os.listdir(category_path) if os.path.isdir(os.path.join(category_path, f))]
                if program_name in programs:
                    return category
            return None
        except Exception as e:
            handle_exception(e)

    def onInstallFinished(self, program_name, success):
        try:
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
        except Exception as e:
            handle_exception(e)

except Exception as e:
    handle_exception(e)