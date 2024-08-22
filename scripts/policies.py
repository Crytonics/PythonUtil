import json
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
import logging
from PyQt5.QtWidgets import QMessageBox

try:
    # Configure logging
    logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def applyPolicies(self):
        reply = QMessageBox.question(self, 'Confirm', 'Are you sure you want to apply policies?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            logging.info("Applying policies")
            print("Applying policies")
            action = 'apply'
            executePolicies(action)
            QMessageBox.information(None, 'Policies Applied', 'Policies have been successfully applied.')
            logging.info("Policies applied successfully")
            print("Policies applied successfully")

    def revertPolicies(self):
        reply = QMessageBox.question(self, 'Confirm', 'Are you sure you want to revert policies?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            logging.info("Reverting policies")
            print("Reverting policies")
            action = 'revert'
            executePolicies(action)
            QMessageBox.information(None, 'Policies Reverted', 'Policies have been successfully reverted.')
            logging.info("Policies reverted successfully")
            print("Policies reverted successfully")

    def executePolicies(action):
        # Load policies from JSON file
        with open('functions/policies/policies.json', 'r') as file:
            policies = json.load(file)

        def set_policy(policy):
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
    logging.error("An error occurred: %s", e)
    print("An error occurred: %s", e)
    QMessageBox.critical(None, "Error", f"An error occurred: {e}")