import pandas as pd
import os

def listrevision_txt_files(self):
    entry1_value = self.entry1.get()    
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    directory_path = os.path.join(initial_dir, entry1_value,"FOD", 'Report')
    print(f"Directory path: {directory_path}")
    if not os.path.isdir(directory_path):
        print(f"Directory does not exist: {directory_path}")
        return
    txt_file_names = [
        filename for filename in os.listdir(directory_path)
        if filename.lower().endswith(".txt") and os.path.getsize(os.path.join(directory_path, filename)) > 0
    ]
    print(f"Found non-empty .txt files: {txt_file_names}")
    os.makedirs(directory_path, exist_ok=True)
    output_file = os.path.join(directory_path, 'RevisedForms.txt')
    report_lines = ["List of non-empty .txt files found:"]
    report_lines.extend(txt_file_names)
    with open(output_file, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print(f"Report saved to {output_file}")

def revised2xlsx(self):
    try:
        txtfile = self.entry1.get()  # Assumes this is inside a class with entry1 as a Tkinter Entry widget
        initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
        base_dir = os.path.join(initial_dir, txtfile,"FOD")
        input_txt_path = os.path.join(base_dir, 'Report', 'RevisedForms.txt')
        output_excel_path = os.path.join(base_dir, 'Report', 'RevisedForms.xlsx')

        with open(input_txt_path, "r") as file:
            lines = file.readlines()

        txt_files = [line.strip()[:-4] for line in lines if line.strip().endswith('.txt')]
        df = pd.DataFrame(txt_files, columns=["File Name"])
        df.to_excel(output_excel_path, index=False)

        print(f"Excel file '{output_excel_path}' has been created successfully.")
    except Exception as e:
        print(f"Error in revised2xlsx: {e}")

def match_form_titles(self):
    try:
        txtfile = self.entry1.get()
        initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
        base_dir = os.path.join(initial_dir, txtfile,"FOD")
        revised_excel_path = os.path.join(base_dir, 'Report', 'RevisedForms.xlsx')
        current_formlist_path = os.path.join(base_dir, 'Report', 'Current_FORMLIST.xlsx')
        output_path = os.path.join(base_dir, 'Report', 'Matched_Form_TitlesRevised.xlsx')

        current_forms_df = pd.read_excel(current_formlist_path, sheet_name="Form List", engine="openpyxl")
        revised_forms_df = pd.read_excel(revised_excel_path, sheet_name="Sheet1", engine="openpyxl")

        form_numbers = revised_forms_df["File Name"].dropna().tolist()
        # form_numbers = revised_forms_df["File Name"].tolist()
        # form_numbers_filtered = [form for form in form_numbers if form.startswith("Form")]

        matched_forms_df = current_forms_df[current_forms_df["Form Number"].isin(form_numbers)].copy()
        # matched_forms_df = current_forms_df[current_forms_df["Form Number"].isin(form_numbers_filtered)].copy()

        if 'File Name' in matched_forms_df.columns:
            matched_forms_df.drop(columns=['File Name'], inplace=True)

        matched_forms_df.to_excel(output_path, index=False)
        print(f"Matched form titles have been saved to '{output_path}' without the 'File Name' column.")
    except Exception as e:
        print(f"Error in match_form_titles: {e}")


def match_form_titles1(revised_excel_path, current_formlist_path, output_path):
    try:
        # Load the Excel files
        current_forms_df = pd.read_excel(current_formlist_path, sheet_name="Form List", engine="openpyxl")
        revised_forms_df = pd.read_excel(revised_excel_path, sheet_name="Sheet1", engine="openpyxl")

        # Extract form numbers from the revised forms
        form_numbers = revised_forms_df["File Name"].dropna().tolist()

        # Match form numbers in the current form list
        matched_forms_df = current_forms_df[current_forms_df["Form Number"].isin(form_numbers)].copy()

        # Drop 'File Name' column if it exists
        if 'File Name' in matched_forms_df.columns:
            matched_forms_df.drop(columns=['File Name'], inplace=True)

        # Save the matched forms to a new Excel file
        matched_forms_df.to_excel(output_path, index=False)
        print(f"Matched form titles have been saved to '{output_path}' without the 'File Name' column.")
    except Exception as e:
        print(f"Error in match_form_titles: {e}")

if __name__ == "__main__":
    revised_excel_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\RevisedForms.xlsx"
    current_formlist_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\Current_FORMLIST.xlsx"
    matched_output_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\Matched_Form_Titles.xlsx"

    match_form_titles1(revised_excel_path, current_formlist_path, matched_output_path)

# if __name__ == "__main__":
#     input_txt_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\RevisedForms.txt"
#     revised_excel_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\RevisedForms.xlsx"
#     current_formlist_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\Current_FORMLIST.xlsx"
#     matched_output_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\00081\Report\Matched_Form_Titles.xlsx"

# #     revised2xlsx(input_txt_path, revised_excel_path)
#     match_form_titles(revised_excel_path, current_formlist_path, matched_output_path)
