import json
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor
import logging

# Configure logging
logging.basicConfig(filename='log.txt', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Check if a command-line argument is provided
if len(sys.argv) < 2:
    print("Usage: policies.py <action>")
    logging.info("Usage: policies.py <action>")
    print("Usage: policies.py <action>")
    sys.exit(1)

action = sys.argv[1]

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

# Exit the script
sys.exit()