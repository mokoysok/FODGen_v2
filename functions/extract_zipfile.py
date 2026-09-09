import os
import zipfile


def select_zip_file(self):
        file_path = self.entry5.get()
        return file_path
    
def extract_zip(zip_file, extract_dir):
        with zipfile.ZipFile(zip_file, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
    
def extract_zip_action(self):
        zip_file = self.entry5.get()
        if not zip_file:
            print("No .zip file selected. Exiting.")
            return
        # Get the directory of the zip file
        zip_dir = os.path.dirname(zip_file)
        # Get the zip file name without extension
        zip_name = os.path.splitext(os.path.basename(zip_file))[0]
        # Automatically create the extraction directory using the zip file name in the same directory as the zip file
        extract_dir = os.path.join(zip_dir, zip_name)
        os.makedirs(extract_dir, exist_ok=True)
        extract_zip(zip_file, extract_dir)
        print(f"Extracted '{zip_file}' to '{extract_dir}' successfully.")
