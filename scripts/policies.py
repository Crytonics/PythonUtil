import json
import subprocess
import sys
import os
from concurrent.futures import ThreadPoolExecutor
import logging
from PyQt5.QtWidgets import QMessageBox
import pkg_resources

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
    def applyPolicies(self):
        try:
            reply = QMessageBox.question(self, 'Confirm', 'Are you sure you want to apply policies?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                logging.info("Applying policies")
                print("Applying policies")
                action = 'apply'
                executePolicies(action)
                QMessageBox.information(None, 'Policies Applied', 'Policies have been successfully applied.')
                logging.info("Policies applied successfully")
                print("Policies applied successfully")
        except Exception as e:
            handle_exception(e)

    def revertPolicies(self):
        try:
            reply = QMessageBox.question(self, 'Confirm', 'Are you sure you want to revert policies?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                logging.info("Reverting policies")
                print("Reverting policies")
                action = 'revert'
                executePolicies(action)
                QMessageBox.information(None, 'Policies Reverted', 'Policies have been successfully reverted.')
                logging.info("Policies reverted successfully")
                print("Policies reverted successfully")
        except Exception as e:
            handle_exception(e)

    def executePolicies(action):
        try:
            # Load policies from JSON file
            with open('functions/policies/policies.json', 'r') as file:
                policies = json.load(file)

            def set_policy(policy):
                try:
                    reg_path = policy['regPath']
                    reg_name = policy['regName']
                    if action == 'apply':
                        reg_value = policy['regValue']
                        type = policy['type']
                        command = f"Set-ItemProperty -Path '{reg_path}' -Name '{reg_name}' -Value '{reg_value}' -Type {type} -Force"
                        logging.info(f"Setting policy: Path='{reg_path}', Name='{reg_name}', Value='{reg_value}' ({type})")
                        print(f"Setting policy: Path='{reg_path}', Name='{reg_name}', Value='{reg_value}' ({type})")
                        subprocess.run(["powershell", "-Command", command], check=True)
                    else:
                        # Check if the item exists before attempting to remove it
                        check_command = f"Test-Path -Path '{reg_path}' -Name '{reg_name}'"
                        result = subprocess.run(["powershell", "-Command", check_command], capture_output=True, text=True)
                        if result.stdout.strip() == 'True':
                            command = f"Remove-ItemProperty -Path '{reg_path}' -Name '{reg_name}' -Force"
                            logging.info(f"Removing policy: Path='{reg_path}', Name='{reg_name}'")
                            print(f"Removing policy: Path='{reg_path}', Name='{reg_name}'")
                            subprocess.run(["powershell", "-Command", command], check=True)
                        else:
                            logging.info(f"Policy not found: Path='{reg_path}', Name='{reg_name}'")
                            print(f"Policy not found: Path='{reg_path}', Name='{reg_name}'")
                except Exception as e:
                    handle_exception(e)

            # Apply each policy using multithreading
            with ThreadPoolExecutor() as executor:
                executor.map(set_policy, policies)

            # Refresh group policy settings
            logging.info("Refreshing group policy settings...")
            print("Refreshing group policy settings...")
            subprocess.run(["gpupdate", "/force"], check=True)
            logging.info("Finished applying policies.")
            print("Finished applying policies.")
        except Exception as e:
            handle_exception(e)

    def installPythonModules(self):
        try:
            logging.info("Starting installation of Python modules.")
            print("Starting installation of Python modules.")
            
            with open('functions/python/modules.txt', 'r') as file:
                requirements = file.read().splitlines()
                
                installed_packages = {pkg.key for pkg in pkg_resources.working_set}
                missing_packages = [pkg for pkg in requirements if pkg not in installed_packages]

                if missing_packages:
                    subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing_packages])
            
            logging.info("Finished installation of Python modules.")
            print("Finished installation of Python modules.")

        except Exception as e:  
            handle_exception(e)

except Exception as e:
    handle_exception(e)