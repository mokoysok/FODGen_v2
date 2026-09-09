import os
import zipfile
import sys

def zip_rtf_files(self):
    # Retrieve the folder name from self.entry1.
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry

    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # folder_path = os.path.join(initial_dir, folder_entry)    
    folder_path = os.path.join(initial_dir, folder_entry, "FOD")

    # Construct the desired zip file name using the folder entry and a release number from entry2.
    zip_filename = f"{temp_entry.get()}_{self.entry2.get()}.zip"
    zip_file_path = os.path.join(folder_path, zip_filename)

    # Validate if the provided folder path exists.
    if not os.path.isdir(folder_path):
        print(f"Error: The folder '{folder_path}' does not exist or is not a directory.")
        return

    rtf_files_found = []
    # Iterate through the folder to find all .rtf files.
    for item in os.listdir(folder_path):
        item_path = os.path.join(folder_path, item)
        if os.path.isfile(item_path) and item.lower().endswith('.rtf'):
            rtf_files_found.append(item_path)

    if not rtf_files_found:
        print(f"No .rtf files found in '{folder_path}'. No zip file will be created.")
        return

    print(f"Found {len(rtf_files_found)} .rtf file(s) in '{folder_path}'.")
    print(f"Creating zip file: '{zip_file_path}'...")

    try:
        # Create a zip archive.
        with zipfile.ZipFile(zip_file_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in rtf_files_found:
                # Add the file to the zip archive.
                # arcname ensures the file is stored with just its name in the zip.
                zipf.write(file_path, arcname=os.path.basename(file_path))
                print(f"  Added: {os.path.basename(file_path)}")

        print(f"Successfully created '{zip_filename}' containing all .rtf files in '{folder_path}'.")

    except Exception as e:
        print(f"An error occurred during zipping: {e}")
        # Clean up a partially created zip file if an error occurs.
        if os.path.exists(zip_file_path):
            os.remove(zip_file_path)
            print(f"Cleaned up incomplete zip file: {zip_file_path}")
