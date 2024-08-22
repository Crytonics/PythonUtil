import json
import subprocess
import logging
from PyQt5.QtWidgets import QListWidgetItem, QMessageBox
from PyQt5.QtCore import Qt

# Configure logging
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

try:
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

except Exception as e:
    logging.error("An error occurred: %s", e)
    print("An error occurred: %s", e)
    QMessageBox.critical(None, "Error", f"An error occurred: {e}")