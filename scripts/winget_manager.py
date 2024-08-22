import json
import subprocess
import logging
import os
import sys
from PyQt5.QtWidgets import QListWidgetItem, QMessageBox
from PyQt5.QtCore import Qt

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
    sys.exit(1)

try:
    def loadWingetData(self):
        try:
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
                    
                    # Create a new category item for the second list
                    category_item_copy = QListWidgetItem(category)
                    category_item_copy.setFlags(category_item.flags() & ~Qt.ItemIsUserCheckable)
                    category_item_copy.setBackground(Qt.lightGray)
                    self.listWidgetWingetUpdaUnins.addItem(category_item_copy)
                    
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
                        
                        # Create a new list item for the second list
                        list_item_copy = QListWidgetItem(details['Name'])
                        list_item_copy.setFlags(list_item.flags() | Qt.ItemIsUserCheckable)
                        list_item_copy.setCheckState(Qt.Unchecked)
                        list_item_copy.setData(Qt.UserRole, details['Name'])
                        if self.is_program_installed_winget(details['Name']): 
                            list_item_copy.setFlags(list_item.flags() | Qt.ItemIsEnabled)
                            list_item_copy.setText(f"{details['Name']} (Installed)")
                        else:
                            list_item_copy.setBackground(Qt.lightGray)
                            list_item_copy.setFlags(list_item.flags() & ~Qt.ItemIsEnabled)
                        self.listWidgetWingetUpdaUnins.addItem(list_item_copy)
        except Exception as e:
            handle_exception(e)

    def is_program_installed_winget(self, program_name):
        try:
            if not hasattr(self, 'installed_programs'):
                print("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                logging.info("***[Checking Installed Programs]*** Fetching list of installed programs using winget")
                result = subprocess.run(["winget", "list"], capture_output=True, text=True, check=True)
                self.installed_programs = result.stdout.lower()
            
            if program_name.lower() in self.installed_programs:
                print(f"***[Checking Installed Programs]*** {program_name} is installed")
                logging.info(f"***[Checking Installed Programs]*** {program_name} is installed")
            else:
                print(f"***[Checking Installed Programs]*** {program_name} is not installed")
                logging.info(f"***[Checking Installed Programs]*** {program_name} is not installed")
            return program_name.lower() in self.installed_programs
        except Exception as e:
            handle_exception(e)

    def selectAllWinget(self, state):
        try:
            for i in range(self.listWidgetWinget.count()):
                item = self.listWidgetWinget.item(i)
                if item.background() == Qt.lightGray:  # Check if the item is a category
                    self.categorySelectAllWinget(item, state)
                elif item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)
        except Exception as e:
            handle_exception(e)

    def selectAllWingetUpdaUnins(self, state):
        try:
            for i in range(self.listWidgetWingetUpdaUnins.count()):
                item = self.listWidgetWingetUpdaUnins.item(i)
                if item.background() == Qt.lightGray:  # Check if the item is a category
                    self.categorySelectAllWingetUpdaUnins(item, state)
                elif item.background() == Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)
        except Exception as e:
            handle_exception(e)

    def categorySelectAllWinget(self, category_item, state):
        try:
            category_index = self.listWidgetWinget.row(category_item)
            for i in range(category_index + 1, self.listWidgetWinget.count()):
                item = self.listWidgetWinget.item(i)
                if item.background() != Qt.lightGray:  # Stop if another category is encountered
                    break
                if item.background() == Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)
        except Exception as e:
            handle_exception(e)

    def categorySelectAllWingetUpdaUnins(self, category_item2, state):
        try:
            category_index_2= self.listWidgetWingetUpdaUnins.row(category_item2)
            for i in range(category_index_2 + 1, self.listWidgetWingetUpdaUnins.count()):
                item = self.listWidgetWingetUpdaUnins.item(i)
                if item.background() == Qt.lightGray:  # Stop if another category is encountered
                    break
                if item.background() != Qt.green and item.flags() & Qt.ItemIsEnabled:
                    item.setCheckState(Qt.Checked if state == Qt.Checked else Qt.Unchecked)
        except Exception as e:
            handle_exception(e)

    def onItemChangedWinget(self, item):
        try:
            if item.background() == Qt.lightGray:  # Check if the item is a category
                self.categorySelectAllWinget(item)
        except Exception as e:
            handle_exception(e)

    def onItemChangedWingetUpdaUnins(self, item):
        try:
            if item.background() == Qt.lightGray:  # Check if the item is a category
                self.categorySelectAllWingetUpdaUnins(item)
        except Exception as e:
            handle_exception(e)

    def installSelectedWinget(self):
        try:
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
        except Exception as e:
            handle_exception(e)

    def installNextWinget(self):
        try:
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
        except Exception as e:
            handle_exception(e)

    def onInstallFinishedWinget(self, program_name, success):
        try:
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
        except Exception as e:
            handle_exception(e)

    def updateCounterLabelWinget(self):
        try:
            self.counterLabelWinget.setText(f'Installed: {self.installedCountWinget}/{self.totalProgramsWinget}')
        except Exception as e:
            handle_exception(e)

    
    # ------- Uninstall section -------
    def uninstallSelectedWingetUpdaUnins(self):
        try:
            selected_items = [self.listWidgetWingetUpdaUnins.item(i) for i in range(self.listWidgetWingetUpdaUnins.count()) if self.listWidgetWingetUpdaUnins.item(i).checkState() == Qt.Checked]
            
            # Reset text for items marked as failed
            for i in range(self.listWidgetWingetUpdaUnins.count()):
                item = self.listWidgetWingetUpdaUnins.item(i)
                if "(Installed)" in item.text():
                    item.setText(item.text().replace(" (Installed)", ""))
                elif "(Failed)" in item.text():
                    item.setText(item.text().replace(" (Failed)", ""))
            

            if selected_items:
                self.uninstallQueueWinget = selected_items
                self.progressBarWingetUpdaUnins.setMaximum(100)
                self.progressBarWingetUpdaUnins.setValue(0)
                self.incrementWingetUpdaUnins = 100 / len(selected_items)  # Calculate increment based on the number of selected items
                self.uninstallNextWinget()
            else:
                QMessageBox.warning(self, 'No Selection', 'Please select programs to uninstall')
        except Exception as e:
            handle_exception(e)

    def uninstallNextWinget(self):
        try:
            if self.uninstallQueueWinget:
                item = self.uninstallQueueWinget.pop(0)
                program_name = item.data(Qt.UserRole)  # Retrieve the actual program name for uninstallation
                winget_command = self.winget_data[program_name]['winget']
                ps_command = f"winget uninstall {winget_command}"
                try:
                    logging.info(f"***[Winget]*** Uninstalling {program_name} using winget")
                    print(f"***[Winget]*** Uninstalling {program_name} using winget")
                    subprocess.run(["powershell", "-Command", ps_command], check=True)
                    logging.info(f"***[Winget]*** Successfully uninstalled {program_name}")
                    print(f"***[Winget]*** Successfully uninstalled {program_name}")
                    self.onUninstallFinishedWinget(program_name, True)
                except subprocess.CalledProcessError:
                    logging.error(f"***[Winget]*** Failed to uninstall {program_name}")
                    print(f"***[Winget]*** Failed to uninstall {program_name}")
                    self.onUninstallFinishedWinget(program_name, False)
            else:
                QMessageBox.information(self, 'Uninstall', 'All selected programs have been uninstalled')
                self.progressBarWingetUpdaUnins.setValue(100)  # Ensure progress bar is set to 100% when done
        except Exception as e:
            handle_exception(e)

    def onUninstallFinishedWinget(self, program_name, success):
        try:
            for i in range(self.listWidgetWingetUpdaUnins.count()):
                item = self.listWidgetWingetUpdaUnins.item(i)
                item_text = item.text().strip()  # Remove leading and trailing spaces
                if item_text == self.winget_data[program_name]['Name']:
                    if success:
                        item.setBackground(Qt.grey)
                        item.setCheckState(Qt.Unchecked)
                        item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
                        item.setText(f"{item.text()} (Uninstalled)")
                    else:
                        item.setBackground(Qt.red)
                        item.setText(f"{item.text()} (Failed)")

                    break
            self.progressBarWingetUpdaUnins.setValue(self.progressBarWingetUpdaUnins.value() + int(self.incrementWingetUpdaUnins))
            self.uninstallNextWinget()  # Continue with the next item
        except Exception as e:
            handle_exception(e)


# ------- Update section -------
    def updateSelectedWingetUpdaUnins(self):
        try:
            selected_items = [self.listWidgetWingetUpdaUnins.item(i) for i in range(self.listWidgetWingetUpdaUnins.count()) if self.listWidgetWingetUpdaUnins.item(i).checkState() == Qt.Checked]
            
            # Reset text for items marked as failed
            for i in range(self.listWidgetWingetUpdaUnins.count()):
                item = self.listWidgetWingetUpdaUnins.item(i)
                if "(Installed)" in item.text():
                    item.setText(item.text().replace(" (Installed)", ""))
                elif "(Failed)" in item.text():
                    item.setText(item.text().replace(" (Failed)", ""))
                elif "(Installed) (Update available)" in item.text():
                    item.setText(item.text().replace(" (Installed) (Update available)", ""))
                elif "(Updated)" in item.text():
                    item.setText(item.text().replace(" (Updated)", ""))
                elif "(Failed) (Update available)" in item.text():
                    item.setText(item.text().replace(" (Failed) (Update available)", ""))
            

            if selected_items:
                self.updateQueueWinget = selected_items
                self.progressBarWingetUpdaUnins.setMaximum(100)
                self.progressBarWingetUpdaUnins.setValue(0)
                self.incrementWingetUpdaUnins = 100 / len(selected_items)  # Calculate increment based on the number of selected items
                self.updateNextWinget()
            else:
                QMessageBox.warning(self, 'No Selection', 'Please select programs to update')
        except Exception as e:
            handle_exception(e)

    def updateNextWinget(self):
        try:
            if self.updateQueueWinget:
                item = self.updateQueueWinget.pop(0)
                program_name = item.data(Qt.UserRole)  # Retrieve the actual program name for uninstallation
                winget_command = self.winget_data[program_name]['winget']
                ps_command = f"winget upgrade {winget_command}"
                try:
                    logging.info(f"***[Winget]*** Updating {program_name} using winget")
                    print(f"***[Winget]*** Updating {program_name} using winget")
                    subprocess.run(["powershell", "-Command", ps_command], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=True)
                    logging.info(f"***[Winget]*** Successfully updated {program_name}")
                    print(f"***[Winget]*** Successfully updated {program_name}")
                    self.onUpdateFinishedWinget(program_name, True)
                except subprocess.CalledProcessError as e:
                    if "No newer package versions are available from the configured sources." in e.stdout:
                        logging.info(f"***[Winget]*** {program_name} is already up-to-date")
                        print(f"***[Winget]*** {program_name} is already up-to-date")
                        self.onUpdateFinishedWinget(program_name, True)
                    else:
                        logging.error(f"***[Winget]*** Failed to update {program_name}")
                        print(f"***[Winget]*** Failed to update {program_name}")
                        logging.error(e.stdout)
                        self.onUpdateFinishedWinget(program_name, False)
            else:
                QMessageBox.information(self, 'Update', 'All selected programs have been updated')
                self.progressBarWingetUpdaUnins.setValue(100)  # Ensure progress bar is set to 100% when done
        except Exception as e:
            handle_exception(e)

    def onUpdateFinishedWinget(self, program_name, success):
        try:
            for i in range(self.listWidgetWingetUpdaUnins.count()):
                item = self.listWidgetWingetUpdaUnins.item(i)
                item_text = item.text().strip()  # Remove leading and trailing spaces
                if item_text == self.winget_data[program_name]['Name']:
                    if success:
                        item.setBackground(Qt.green)
                        item.setText(f"{item.text()} (Updated)")
                    else:
                        item.setBackground(Qt.red)
                        item.setText(f"{item.text()} (Failed)")

                    break
            self.progressBarWingetUpdaUnins.setValue(self.progressBarWingetUpdaUnins.value() + int(self.incrementWingetUpdaUnins))
            self.updateNextWinget()  # Continue with the next item
        except Exception as e:
            handle_exception(e)

except Exception as e:
    handle_exception(e)