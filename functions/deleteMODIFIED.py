import os
import re

def delete_modified_userguide_rtf(self):
    temp_entry = self.entry1() if callable(self.entry1) else self.entry1
    folder_entry = temp_entry.get() if hasattr(temp_entry, "get") else temp_entry

    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")    
    pattern = re.compile(r"(.+)userguide_MODIFIED\.rtf", re.IGNORECASE) and re.compile(r"userguide_MODIFIED\.rtf", re.IGNORECASE)
    if not os.path.isdir(base_dir):
        print(f"Error: The folder '{base_dir}' does not exist or is not a directory.")
        return

    deleted_count = 0

    for item in os.listdir(base_dir):
        item_path = os.path.join(base_dir, item)

        if os.path.isfile(item_path) and pattern.match(item):
            try:
                os.remove(item_path)
                print(f"  Deleted: {item}")
                deleted_count += 1
            except OSError as e:
                print(f"  Error deleting '{item}': {e}")
            except Exception as e:
                print(f"  An unexpected error occurred with '{item}': {e}")

    if deleted_count > 0:
        print(f"\nSuccessfully deleted {deleted_count} file(s) matching the pattern.")
    else:
        print("\nNo files matching the pattern were found and deleted.")