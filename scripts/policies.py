import json
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor

# Load policies from JSON file
with open('functions/policies/policies.json', 'r') as file:
    policies = json.load(file)

# Function to set a registry policy using PowerShell
def set_policy(policy):
    reg_path = policy['regPath']
    reg_name = policy['regName']
    reg_value = policy['regValue']
    type = policy['type']
    # Ensure the registry key exists and set the property with the specified type
    command = f"Set-ItemProperty -Path '{reg_path}' -Name '{reg_name}' -Value '{reg_value}' -Type {type} -Force"
    print(f"Setting policy: Path='{reg_path}', Name='{reg_name}', Value='{reg_value}' ({type})")
    subprocess.run(["powershell", "-Command", command], check=True)

# Apply each policy using multithreading
with ThreadPoolExecutor() as executor:
    executor.map(set_policy, policies)

# Refresh group policy settings
print("Refreshing group policy settings...")
subprocess.run(["gpupdate", "/force"], check=True)
print("Finished applying policies.")

# Exit the script
sys.exit()