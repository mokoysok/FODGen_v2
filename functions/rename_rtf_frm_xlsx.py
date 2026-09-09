import os
import shutil
import pandas as pd

def rename_rtf_from_excel1(self):
    txtfile = self.entry1.get()
    initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary"
    base_dir = os.path.join(initial_dir, txtfile)
    target_folder = os.path.join(base_dir, 'FOD')
    report_dir = os.path.join(target_folder, 'Report')

    # Ensure report and modified folders exist
    os.makedirs(report_dir, exist_ok=True)
    modified_folder = os.path.join(target_folder, 'MODIFIED')
    os.makedirs(modified_folder, exist_ok=True)

    # Load the Excel file
    excel_file = os.path.join(report_dir, 'Current_FORMLIST.xlsx')
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"Excel file '{excel_file}' not found.")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Create a dictionary mapping original file names to form numbers
    if 'File Name' not in df.columns or 'Form Number' not in df.columns:
        print("Excel file must contain 'File Name' and 'Form Number' columns.")
        return

    file_mapping = dict(zip(df['File Name'], df['Form Number']))

    # Initialize counters
    renamed_count = 0
    total_files_in_excel = len(file_mapping)
    unmatched_files = []

    # Iterate through the files in the folder and rename them based on the mapping
    for filename in os.listdir(target_folder):
        if filename.endswith('.rtf'):
            match_key = next((key for key in file_mapping if key.lower() == filename.lower()), None)
            if match_key:
                old_path = os.path.join(target_folder, filename)
                new_filename = f"{file_mapping[match_key]}.rtf"
                new_path = os.path.join(modified_folder, new_filename)
                shutil.copy2(old_path, new_path)
                renamed_count += 1
            else:
                unmatched_files.append(filename)

    # Print report
    print("Renaming and copying completed successfully.")
    print(f"Total .rtf files listed in Current FormList Excel: {total_files_in_excel}")
    print(f"Number of .rtf files successfully renamed and copied: {renamed_count}")
    print(f"Number of .rtf files not matched: {len(unmatched_files)}")
    if unmatched_files:
        print("Unmatched files:")
        for f in unmatched_files:
            print(f" - {f}")

def rename_rtf_from_excel2(self):
    # Get folder name from user input
    txtfile = self.entry1.get()
    prevpub = self.entry5.get().strip().replace('.zip', '')

    # Define base and report directories
    initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary"
    base_dir = os.path.join(initial_dir, txtfile, "FOD")
    target_folder = os.path.join(base_dir, prevpub)
    report_dir = os.path.join(base_dir, 'Report')

    # Ensure report and modified folders exist
    os.makedirs(report_dir, exist_ok=True)
    modified_folder = os.path.join(target_folder, 'MODIFIED')
    os.makedirs(modified_folder, exist_ok=True)

    # Load the Excel file
    excel_file = os.path.join(report_dir, 'Previous_FORMLIST.xlsx')
    try:
        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        print(f"Excel file '{excel_file}' not found.")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    # Create a dictionary mapping original file names to form numbers
    if 'File Name' not in df.columns or 'Form Number' not in df.columns:
        print("Excel file must contain 'File Name' and 'Form Number' columns.")
        return

    file_mapping = dict(zip(df['File Name'], df['Form Number']))

    # Initialize counters
    renamed_count = 0
    total_files_in_excel = len(file_mapping)
    unmatched_files = []

    # Iterate through the files in the folder and rename them based on the mapping
    for filename in os.listdir(target_folder):
        if filename.endswith('.rtf'):
            match_key = next((key for key in file_mapping if key.lower() == filename.lower()), None)
            if match_key:
                old_path = os.path.join(target_folder, filename)
                new_filename = f"{file_mapping[match_key]}.rtf"
                new_path = os.path.join(modified_folder, new_filename)
                shutil.copy2(old_path, new_path)
                renamed_count += 1
            else:
                unmatched_files.append(filename)

    # Print report
    print("Renaming and copying completed successfully.")
    print(f"Total .rtf files listed in Previous Form List Excel: {total_files_in_excel}")
    print(f"Number of .rtf files successfully renamed and copied: {renamed_count}")
    print(f"Number of .rtf files not matched: {len(unmatched_files)}")
    if unmatched_files:
        print("Unmatched files:")
        for f in unmatched_files:
            print(f" - {f}")            

# if __name__ == "__main__":
#     rename_rtf_from_excel(
#         excel_file=r'C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\test\formlist\form_listtxt2xlsx.xlsx',
#         target_folder=r'C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\test\formlist'
#     )