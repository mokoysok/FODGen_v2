import os
import shutil

def clean_directories(self):
    # Retrieve the folder name from self.entry1.
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry
    
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    zip_file = self.entry5.get().strip()  # get ZIP file path from self.entry5
    if zip_file:
        zip_dir = os.path.dirname(zip_file)
        zip_name = os.path.splitext(os.path.basename(zip_file))[0]
        extract_dir = os.path.join(zip_dir, zip_name)
    else:
        extract_dir = None

    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_directory = os.path.join(base_dir, 'Report')
    
    # Delete extract_dir if it exists.
    if extract_dir and os.path.exists(extract_dir):
        try:
            shutil.rmtree(extract_dir)
            print(f"Deleted extract directory: {extract_dir}")
        except Exception as e:
            print(f"Error deleting extract directory '{extract_dir}': {e}")
    else:
        print("Extract directory does not exist or ZIP file not specified.")
    
    # Delete report_directory if it exists.
    if os.path.exists(report_directory):
        try:
            shutil.rmtree(report_directory)
            print(f"Deleted report directory: {report_directory}")
        except Exception as e:
            print(f"Error deleting report directory '{report_directory}': {e}")
    else:
        print("Report directory does not exist.")
