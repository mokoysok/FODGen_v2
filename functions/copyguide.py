import shutil
import os

def copy_file(self):
    # Retrieve the folder name from self.entry1
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry

    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    # source_directory = r"C:\Users\LABRADBM\Downloads\Local\eComp Project\FOD"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    source_directory = r"\\fabwebd5.net\neptune\DataConversion\Tools\FOD\reference"    
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_directory = os.path.join(base_dir, 'Report')
    
    os.makedirs(report_directory, exist_ok=True)
    filename = "ref_userguide.rtf"
    source_folder = source_directory         
    destination_folder = report_directory     
    source_path = os.path.join(source_folder, filename)
    destination_path = os.path.join(destination_folder, filename)

    if not os.path.exists(source_path):
        print(f"Error: Source file '{source_path}' does not exist.")
        return

    os.makedirs(destination_folder, exist_ok=True)
    try:
        shutil.copy2(source_path, destination_path)
        print(f"File '{filename}' copied successfully from '{source_folder}' to '{destination_folder}'.")
    except shutil.Error as e:
        print(f"Error copying file: {e}")
    except IOError as e:
        print(f"I/O error: {e}")
