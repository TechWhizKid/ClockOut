import winreg
import ctypes
import sys
import os

# Check if script has admin permits or not
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Create registry key
def create_registry_key(file_name):
    # Get the application path using the current script path
    script_path = os.path.abspath(sys.argv[0])
    script_directory = os.path.dirname(script_path)

    # Form the file path for the registry key using the script directory and the passed file name
    key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    file_path = os.path.join(script_directory, file_name)

    # Open the desired registry key for auto-start programs
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)

    # Set the registry key with the file path
    winreg.SetValueEx(key, file_name, 0, winreg.REG_SZ, file_path)
    print(f"The registry key for '{file_name}' was added successfully.")

    # Close the registry key
    winreg.CloseKey(key)

if __name__ == "__main__":
    if len(sys.argv) == 2:
        file_name = sys.argv[1]

        if is_admin():
            create_registry_key(file_name)
        else:
            # Set the script window to stay on top
            hwnd = ctypes.windll.kernel32.GetConsoleWindow()
            if hwnd != 0:
                ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0020)
            
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    else:
        print(f"Usage: {os.path.basename(__file__)} <file_name>")
