import json
import winreg
import subprocess
import sys
import logging
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QListWidgetItem, QMessageBox
from PyQt5.QtCore import Qt

# Configure logging
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def get_uninstall_command(program_name):
    uninstall_keys = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    
    for uninstall_key in uninstall_keys:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, uninstall_key) as key:
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                sub_key_name = winreg.EnumKey(key, i)
                with winreg.OpenKey(key, sub_key_name) as sub_key:
                    try:
                        display_name = winreg.QueryValueEx(sub_key, "DisplayName")[0]
                        if program_name.lower() in display_name.lower():
                            return winreg.QueryValueEx(sub_key, "UninstallString")[0]
                    except FileNotFoundError:
                        pass
    return None

def uninstall_appx_package(program_name):
    ps_command = f"Get-AppxPackage *{program_name}* | Remove-AppxPackage"
    result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True, shell=True)
    return result.returncode == 0 and not result.stderr.strip()

def uninstall_program(program_name, is_appx_package):
    if is_appx_package:
        success = uninstall_appx_package(program_name)
        if success:
            logging.info(f"{program_name} uninstalled successfully.")
            print(f"{program_name} uninstalled successfully.")
        else:
            logging.info(f"Failed to uninstall {program_name}.")
            print(f"Failed to uninstall {program_name}.")
        return success
    else:
        uninstall_command = get_uninstall_command(program_name)
        if uninstall_command:
            try:
                subprocess.run(uninstall_command, shell=True, check=True)
                logging.info(f"{program_name} uninstalled successfully.")
                print(f"{program_name} uninstalled successfully.")
                return True
            except subprocess.CalledProcessError:
                logging.info(f"Failed to uninstall {program_name}.")
                print(f"Failed to uninstall {program_name}.")
                return False
        else:
            logging.info(f"Uninstall command for {program_name} not found.")
            print(f"Uninstall command for {program_name} not found.")
            return False

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
            success = uninstall_program(self.program_name, self.is_appx_package)
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

def is_appx_package_installed(app_name):
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
            if not self.is_program_installed_for_uninstall(item['name']) and not is_appx_package_installed(item['name']):
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
        is_appx_package = is_appx_package_installed(program_name)
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

def main():
    if len(sys.argv) > 2:
        program_name = sys.argv[1]
        is_appx_package = sys.argv[2].lower() == 'true'
        success = uninstall_program(program_name, is_appx_package)
        sys.exit(0 if success else 1)
    else:
        logging.info("No program name or package type provided.")
        print("No program name or package type provided.")
        sys.exit(1)

if __name__ == "__main__":
    main()