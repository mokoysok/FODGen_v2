import shutil
import os
import re

def move_rtf_files(self):
    # Retrieve the folder name from self.entry1.
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry

    # Define the initial directory and build the base directory from the folder entry.
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    # base_dir = os.path.join(initial_dir, folder_entry)
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    
    # Define the destination "Old" folder inside base_dir.
    old_folder = os.path.join(base_dir, 'Old')
    os.makedirs(old_folder, exist_ok=True)
    
    # Iterate over all files in base_dir.
    for filename in os.listdir(base_dir):
        file_path = os.path.join(base_dir, filename)
        # Skip if it's a directory
        if not os.path.isfile(file_path):
            continue


        # Define the pattern to exclude user guide files (case-insensitive)
        exclude_pattern = re.compile(r"(.*)?user\s?guide\.rtf$", re.IGNORECASE)

        if filename.lower().endswith(".rtf") and not exclude_pattern.match(filename):
            # Destination file retains its original filename
            destination_path = os.path.join(old_folder, filename)
            try:
                shutil.move(file_path, destination_path)
                print(f"Moved rtf file '{filename}' to '{old_folder}'.")
            except shutil.Error as e:
                print(f"Error moving file '{filename}': {e}")
            except IOError as e:
                print(f"I/O error moving file '{filename}': {e}")
        #         # Check if it's an .rtf file (case insensitive) and ensure the file is not "userguide.rtf"
        # if filename.lower().endswith(".rtf") and not re.match(r"(.+)userguide\.rtf", filename.lower()) and not re.match(r"userguide\.rtf", filename.lower()) and not re.match(r"User\sGuide\.rtf", filename.lower()) and not re.match(r"UserGuide\.rtf", filename.lower()) and not re.match(r"(.+)User\sGuide\.rtf", filename.lower()) and not re.match(r"(.+)UserGuide\.rtf", filename.lower()):
        #     # Destination file retains its original filename
        #     destination_path = os.path.join(old_folder, filename)
        #     try:
        #         shutil.move(file_path, destination_path)
        #         print(f"Moved '{filename}' to '{old_folder}'.")
        #     except shutil.Error as e:
        #         print(f"Error moving file '{filename}': {e}")
        #     except IOError as e:
        #         print(f"I/O error moving file '{filename}': {e}")

# For testing, a dummy class simulating a GUI widget with an entry1 method.
# if __name__ == "__main__":
#     class Dummy:
#         def __init__(self, folder):
#             self.entry1 = lambda: DummyEntry(folder)
    
#     class DummyEntry:
#         def __init__(self, text):
#             self.text = text
#         def get(self):
#             return self.text

#     # Replace "00493" with your desired folder name.
#     dummy_self = Dummy("00491")
#     move_rtf_files(dummy_self)
