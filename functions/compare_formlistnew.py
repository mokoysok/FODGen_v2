import pandas as pd
import os

def compare_formlists(self):
    txtfile = self.entry1.get()
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    base_dir = os.path.join(initial_dir, txtfile,"FOD")
    previous_excel_path = os.path.join(base_dir, 'Report', 'Previous_FORMLIST.xlsx')
    current_excel_path = os.path.join(base_dir, 'Report', 'Current_FORMLIST.xlsx')
    matched_output_path_add = os.path.join(base_dir, 'Report', 'AddedForms.xlsx')
    matched_output_path_del = os.path.join(base_dir, 'Report', 'DeletedForms.xlsx')
    # Load the Excel files using the provided paths
    previous_df = pd.read_excel(previous_excel_path, sheet_name="Form List", engine="openpyxl")
    current_df = pd.read_excel(current_excel_path, sheet_name="Form List", engine="openpyxl")

    # Select only the Form Number and Form Title columns
    previous_forms = previous_df[['Form Number']]
    current_forms = current_df[['Form Number']]

    # Merge and identify differences
    comparison = pd.merge(current_forms, previous_forms, on=['Form Number'], how='outer', indicator=True)
    added_only = comparison[comparison['_merge'] == 'left_only'].drop(columns=['_merge'])
    deleted_only = comparison[comparison['_merge'] == 'right_only'].drop(columns=['_merge'])

    # Save the results to Excel files
    added_only.to_excel(matched_output_path_add, index=False)
    deleted_only.to_excel(matched_output_path_del, index=False)
    print(f"Added forms saved to: {matched_output_path_add}")
    print(f"Deleted forms saved to: {matched_output_path_del}")

def match_form_titles_add(self):
    try:
        txtfile = self.entry1.get()
        initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
        base_dir = os.path.join(initial_dir, txtfile,"FOD")
        revised_excel_path = os.path.join(base_dir, 'Report', 'AddedForms.xlsx')
        current_formlist_path = os.path.join(base_dir, 'Report', 'Current_FORMLIST.xlsx')
        output_path = os.path.join(base_dir, 'Report', 'Matched_Form_TitlesAdded.xlsx')

        current_forms_df = pd.read_excel(current_formlist_path, sheet_name="Form List", engine="openpyxl")
        revised_forms_df = pd.read_excel(revised_excel_path, sheet_name="Sheet1", engine="openpyxl")

        # Extract form numbers from revised file
        form_numbers = revised_forms_df["Form Number"].dropna().tolist()

        # Match forms in current list based on Form Number
        matched_forms_df = current_forms_df[current_forms_df["Form Number"].isin(form_numbers)].copy()

        # Drop 'File Name' column if it exists
        if 'File Name' in matched_forms_df.columns:
            matched_forms_df.drop(columns=['File Name'], inplace=True)

        # Save to new Excel file
        matched_forms_df.to_excel(output_path, index=False)
        print(f"Matched form titles have been saved to '{output_path}' without the 'File Name' column.")
    except Exception as e:
        print(f"Error in match_form_titles_deleted: {e}")  


def match_form_titles_deleted(self):
    try:
        txtfile = self.entry1.get()
        initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
        base_dir = os.path.join(initial_dir, txtfile,"FOD")
        revised_excel_path = os.path.join(base_dir, 'Report', 'DeletedForms.xlsx')
        current_formlist_path = os.path.join(base_dir, 'Report', 'Previous_FORMLIST.xlsx')
        output_path = os.path.join(base_dir, 'Report', 'Matched_Form_TitlesDeleted.xlsx')

        # Load Excel files
        current_forms_df = pd.read_excel(current_formlist_path, sheet_name="Form List", engine="openpyxl")
        revised_forms_df = pd.read_excel(revised_excel_path, sheet_name="Sheet1", engine="openpyxl")

        # Extract form numbers from revised file
        form_numbers = revised_forms_df["Form Number"].dropna().tolist()

        # Match forms in current list based on Form Number
        matched_forms_df = current_forms_df[current_forms_df["Form Number"].isin(form_numbers)].copy()

        # Drop 'File Name' column if it exists
        if 'File Name' in matched_forms_df.columns:
            matched_forms_df.drop(columns=['File Name'], inplace=True)

        # Save to new Excel file
        matched_forms_df.to_excel(output_path, index=False)
        print(f"Matched form titles have been saved to '{output_path}' without the 'File Name' column.")
    except Exception as e:
        print(f"Error in match_form_titles_deleted: {e}")
  

# if __name__ == "__main__":
#     previous_excel_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\80940\Report\Previous_FORMLIST.xlsx"
#     current_excel_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\80940\Report\Current_FORMLIST.xlsx"
#     matched_output_path_add = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\80940\Report\Matched_Form_TitlesAdd.xlsx"
#     matched_output_path_del = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\80940\Report\Matched_Form_TitlesDeleted.xlsx"

#     match_form_titles_add(matched_output_path_add, current_excel_path, matched_output_path_add)