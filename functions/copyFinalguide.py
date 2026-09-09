import shutil
import os

def copy_finalfile(self):
    # Retrieve the folder name from self.entry1; handle whether self.entry1 is callable or not.
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry
    userguide = self.entry6.get().strip()  # New name for the copied file

    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_directory = os.path.join(base_dir, 'Report')
    
    # Ensure the Report directory exists (i.e., the source folder)
    os.makedirs(report_directory, exist_ok=True)

    # Define the source file and destination file details.
    filename = "modified_userguide.rtf"
    source_folder = report_directory     
    destination_folder = base_dir             

    # Construct full file paths.
    source_path = os.path.join(source_folder, filename)
    destination_path = os.path.join(destination_folder, userguide)

    if not os.path.exists(source_path):
        print(f"Error: Source file '{source_path}' does not exist.")
        return

    # Ensure the destination folder exists.
    os.makedirs(destination_folder, exist_ok=True)

    try:
        shutil.copy2(source_path, destination_path)
        print(f"File '{filename}' copied successfully from '{source_folder}' to '{destination_path}'.")
    except shutil.Error as e:
        print(f"Error copying file: {e}")
    except IOError as e:
        print(f"I/O error: {e}")