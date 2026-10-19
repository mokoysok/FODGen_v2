import re
import os
from striprtf.striprtf import rtf_to_text

def extract_forms_title_from_rtf(self):
    # Get the folder entry based on the widget type (callable or a Tkinter variable)
    folder_entry = self.entry1() if callable(self.entry1) else self.entry1.get()
    userguide = self.entry6.get().strip()
    
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_directory = os.path.join(base_dir, 'Report')
    
    os.makedirs(report_directory, exist_ok=True)
    file_path = os.path.join(report_directory, userguide)

    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            rtf_content = file.read()
        plain_text = rtf_to_text(rtf_content)

        match = re.search(r'Forms on Download for (.+)', plain_text)
        if match:
            title = match.group(1).strip()
            print(f"Extracted Title: {title}")
            report_path = os.path.join(report_directory, "PubTitle.txt")
            with open(report_path, "w", encoding="utf-8") as report_file:
                report_file.write(f"Extracted Title: {title}\n")

            return title
        else:
            print("No match found.")
            return None

    except Exception as e:
        print(f"Error: {e}")
        return None

def extract_forms_sample_from_rtf(self):
    # Get the folder entry based on the widget type (callable or a Tkinter variable)
    folder_entry = self.entry1() if callable(self.entry1) else self.entry1.get()
    userguide = self.entry6.get().strip()
    
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, folder_entry, "FOD")
    report_directory = os.path.join(base_dir, 'Report')
    
    os.makedirs(report_directory, exist_ok=True)
    file_path = os.path.join(report_directory, userguide)

    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            rtf_content = file.read()
        plain_text = rtf_to_text(rtf_content)

        # Support both legacy "Forms on Disc" and current "Forms on Download"
        # wording. \s also accepts whitespace and line breaks produced by RTF conversion.
        match = re.search(
            r'For\s+example,\s*(.+?)\s+Forms\s+on\s+(?:Disc|Download)\s*,',
            plain_text,
            re.IGNORECASE | re.DOTALL,
        )
        if match:
            sample = re.sub(r'\s+', ' ', match.group(1)).strip()
            print(f"Extracted Sample: {sample}")
            report_path = os.path.join(report_directory, "PubSample.txt")
            with open(report_path, "w", encoding="utf-8") as report_file:
                report_file.write(f"Extracted Sample: {sample}\n")

            return sample
        else:
            print("No match found.")
            return None

    except Exception as e:
        print(f"Error: {e}")
        return None
