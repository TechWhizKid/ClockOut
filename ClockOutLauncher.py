import os
import sys

# Get the directory path of the running executable
current_dir = os.path.dirname(os.path.abspath(sys.executable))

# Specify the folder and file names
folder_name = "App"
file_name = "ClockOut.exe"

# Build the path to the file
file_path = os.path.join(current_dir, folder_name, file_name)

# Check if the folder and file exist
if os.path.isdir(os.path.join(current_dir, folder_name)) and os.path.isfile(file_path):
    # Change the current working directory to the specified folder
    os.chdir(os.path.join(current_dir, folder_name))
    
    # Execute the specified file
    os.startfile(file_name)
else:
    print("Folder or file not found.")
