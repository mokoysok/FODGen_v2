import shutil
import os
import re

def copy_ftoutput_rtf_files(self):
    # Retrieve the folder name from self.entry1.
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry

    # Retrieve the FT output directory from self.entry7.
    ftout_dir = self.entry7.get().strip() if hasattr(self.entry7, "get") else self.entry7

    # Define the initial directory and build the base directory from the folder entry.
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"    
    base_dir = os.path.join(initial_dir, folder_entry,"FOD")
    
    # Ensure the base directory exists.
    os.makedirs(base_dir, exist_ok=True)
    
    # Iterate over all files in the ftout_dir.
    for filename in os.listdir(ftout_dir):
        file_path = os.path.join(ftout_dir, filename)
        # Skip if it's not a file.
        if not os.path.isfile(file_path):
            continue

        # Check if the file has a .rtf extension (case insensitive).
        if filename.lower().endswith(".rtf"):
            destination_path = os.path.join(base_dir, filename)
            try:
                shutil.copy(file_path, destination_path)
                print(f"Copied '{filename}' to '{base_dir}'.")
            except shutil.Error as e:
                print(f"Error copying file '{filename}': {e}")
            except IOError as e:
                print(f"I/O error copying file '{filename}': {e}")

# For testing, a dummy class simulating a GUI widget with an entry1 method.
# if __name__ == "__main__":
#     class Dummy:
#         def __init__(self, folder, ftout):
#             self.entry1 = lambda: DummyEntry(folder)
#             self.entry7 = DummyEntry(ftout)
    
#     class DummyEntry:
#         def __init__(self, text):
#             self.text = text
#         def get(self):
#             return self.text

#     # Replace "00493" with your desired folder name.
#     dummy_self = Dummy("00490", r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00493\00493_68")
#     copy_ftoutput_rtf_files(dummy_self)
