import os
import shutil
import win32com.client
from datetime import datetime, timedelta

# Define the folders
initial_folder = "C:/Users/fhulufhelo/Documents/Automation Scripts/Dump Files Uploader/Original"
new_folder = "C:/Users/fhulufhelo/Documents/Automation Scripts/Dump Files Uploader/Final"

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # 6 is the index of the inbox

# Calculate the time 24 hours ago
time_24hrs_ago = datetime.now() - timedelta(days=1)

# Create a filter for messages received in the last 24 hours
filter_query = "[ReceivedTime] >= '" + time_24hrs_ago.strftime("%m/%d/%Y %H:%M %p") + "'"
messages = inbox.Items.Restrict(filter_query)

# Function to generate a new file name if the file already exists
def get_unique_filename(folder, filename):
    base, extension = os.path.splitext(filename)
    counter = 1
    new_filename = filename
    while os.path.exists(os.path.join(folder, new_filename)):
        new_filename = f"{base}_{counter}{extension}"
        counter += 1
    return new_filename

# Process each filtered email
for message in messages:
    try:
        for attachment in message.Attachments:
            if attachment.FileName.endswith('.dmps'):
                # Save the attachment to the initial folder
                initial_filename = get_unique_filename(initial_folder, attachment.FileName)
                initial_path = os.path.join(initial_folder, initial_filename)
                attachment.SaveAsFile(initial_path)
                print(f"Saved attachment to: {initial_path}")  # Log saving

                # Determine a unique name for the file in the new folder
                new_filename = get_unique_filename(new_folder, initial_filename)
                new_path = os.path.join(new_folder, new_filename)
                
                # Copy the file to the new folder
                shutil.copy2(initial_path, new_path)
                print(f"Copied attachment to: {new_path}")  # Log copying

                # Delete the file from the initial folder
                os.remove(initial_path)
                print(f"Deleted attachment from: {initial_path}")  # Log deletion
    except Exception as e:
        print(f"Error processing message: {e}")

print("Files processed successfully.")
