import os
import re
from striprtf.striprtf import rtf_to_text

def convert_rtf_to_txt_current(self):
    pubnumber = self.entry1.get()
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary"
    
    # base_dir = os.path.join(initial_dir, pubnumber)
    base_dir = os.path.join(initial_dir, pubnumber, "FOD")
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)

    pattern = re.compile(r"^\d+\sFORMLIST\.rtf$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(base_dir):
        if pattern.match(filename):
            input_file = os.path.join(base_dir, filename)
            break

    if input_file is None:
        print("No matching file found.")
        return

    output_file = os.path.join(report_dir, "Current_FORMLIST.txt")
    
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            rtf_content = file.read()
        plain_text = rtf_to_text(rtf_content)
        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(plain_text)
        
        print(f'Successfully converted "{input_file}" to "{output_file}"')
    except Exception as e:
        print("Error:", e)

def convert_rtf_to_txt_prev(self):
    pubnumber = self.entry1.get()
    prevpub = self.entry5.get().strip().replace('.zip', '')
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary"
    
    # base_dir = os.path.join(initial_dir, pubnumber, prevpub)
    base_dirpub = os.path.join(initial_dir, pubnumber, "FOD")
    base_dir = os.path.join(base_dirpub, prevpub)    
    report_dir = os.path.join(base_dirpub,'Report')
    os.makedirs(report_dir, exist_ok=True)

    pattern = re.compile(r"^\d+\sFORMLIST\.rtf$", re.IGNORECASE)
    input_file = None
    for filename in os.listdir(base_dir):
        if pattern.match(filename):
            input_file = os.path.join(base_dir, filename)
            break

    if input_file is None:
        print("No matching file found.")
        return

    output_file = os.path.join(report_dir, "Previous_FORMLIST.txt")
    try:
        with open(input_file, 'r', encoding='utf-8') as file:
            rtf_content = file.read()
        plain_text = rtf_to_text(rtf_content)
        with open(output_file, 'w', encoding='utf-8') as file:
            file.write(plain_text)
        print(f'Successfully converted "{input_file}" to "{output_file}"')
    except Exception as e:
        print("Error:", e)        

