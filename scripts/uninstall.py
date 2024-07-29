import json
import winreg
import subprocess
import sys

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
            print(f"{program_name} uninstalled successfully.")
        else:
            print(f"Failed to uninstall {program_name}.")
        return success
    else:
        uninstall_command = get_uninstall_command(program_name)
        if uninstall_command:
            try:
                subprocess.run(uninstall_command, shell=True, check=True)
                print(f"{program_name} uninstalled successfully.")
                return True
            except subprocess.CalledProcessError:
                print(f"Failed to uninstall {program_name}.")
                return False
        else:
            print(f"Uninstall command for {program_name} not found.")
            return False

def main():
    if len(sys.argv) > 2:
        program_name = sys.argv[1]
        is_appx_package = sys.argv[2].lower() == 'true'
        success = uninstall_program(program_name, is_appx_package)
        sys.exit(0 if success else 1)
    else:
        print("No program name or package type provided.")
        sys.exit(1)

if __name__ == "__main__":
    main()