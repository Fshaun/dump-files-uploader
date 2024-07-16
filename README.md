## Outlook Attachments Downloader

This script automates the process of downloading .dmps files from your Outlook inbox, saving them to a specified folder, copying them to a network location, and then deleting the original downloaded files.

Requirements

- Python 3.x
- pywin32 library (for interacting with Outlook)

Setup

1. Install Python and pip: Ensure you have Python 3.x and pip installed on your system.

2. Install pywin32 library: You can install the required library using pip:
pip install pywin32

3. Update folder paths: Modify the script to set the correct initial and network folder paths.
   
initial_folder = "C:/Users/fhulufhelo/Documents/Automation Scripts/Dump Files Uploader/Original"
new_folder = r"\\NetworkLocation\SharedFolder\Dump Files Uploader\Final"

Script Description

The script performs the following tasks:

1. Connects to the Outlook inbox.
2. Filters emails received in the last 24 hours.
3. Downloads attachments with a .dmps file extension to the initial folder.
4. Renames the file if a file with the same name already exists in the initial folder.
5. Copies the file to a specified network location, ensuring no file name conflicts.
6. Deletes the original downloaded file from the initial folder.

Usage

1. Ensure Outlook is running: The script connects to the Outlook application, which needs to be running.

2. Run the script: Execute the script using Python.
python outlook_attachments_downloader.py

3. Check the log output: The script prints messages to the console for each step, indicating the progress and any errors encountered.

Notes

- Ensure the network location is accessible and you have the necessary permissions to read and write files.
- Modify the filter criteria and file extension as needed for different use cases.
  
Troubleshooting
- Permission Errors: Ensure you have read/write permissions for both the initial and network folders.
- Outlook Connection Issues: Ensure Outlook is running and you are logged in with the correct profile.
