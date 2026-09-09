import os
import csv
import openpyxl

# def formlisttxt2csv(input_path, output_path):
#     # Read the file content
#     with open(input_path, "r", encoding="utf-8") as file:
#         lines = file.readlines()

#     entries = []
#     rtf_count = 0

#     # Extract relevant entries
#     for line in lines:
#         line = line.strip()
#         if ".rtf" in line:
#             rtf_count += 1
#         if line and '|' in line and not line.startswith("File Name:"):
#             parts = line.split('|')
#             if len(parts) >= 3:
#                 file_name = parts[0].strip()
#                 form_number = parts[1].strip()
#                 form_title = parts[2].strip()
#                 if file_name and form_number and form_title:
#                     entries.append([file_name, form_number, form_title])

#     # Write to CSV
#     with open(output_path, mode='w', newline='', encoding='utf-8') as f:
#         writer = csv.writer(f)
#         writer.writerow(["File Name", "Form Number", "Form Title"])
#         writer.writerows(entries)

#     # Compare counts
#     print(f"Number of '.rtf' entries in text file: {rtf_count}")
#     print(f"Number of non-empty rows written to CSV: {len(entries)}")
#     print("✅ Counts match." if rtf_count == len(entries) else "❌ Counts do not match.")


def formlisttxt2xlsx_current(self):
    txtfile = self.entry1.get()      
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    base_dir = os.path.join(initial_dir, txtfile, "FOD")
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)

    input_path = os.path.join(report_dir, "Current_FORMLIST.txt")
    output_path = os.path.join(report_dir, "Current_FORMLIST.xlsx")

    # Read the file content
    with open(input_path, "r", encoding="utf-8") as file:
        lines = file.readlines()

    entries = []
    rtf_count = 0

    # Extract relevant entries
    for line in lines:
        line = line.strip()
        if ".rtf" in line:
            rtf_count += 1
        if line and '|' in line and not line.startswith("File Name:"):
            parts = line.split('|')
            if len(parts) >= 3:
                file_name = parts[0].strip()
                form_number = parts[1].strip()
                form_title = parts[2].strip()
                if file_name and form_number and form_title:
                    entries.append([file_name, form_number, form_title])

    # Write to Excel (.xlsx)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Form List"
    ws.append(["File Name", "Form Number", "Form Title"])
    for entry in entries:
        ws.append(entry)
    wb.save(output_path)

    # Compare counts
    print(f"Number of '.rtf' entries in Current Form list txt file: {rtf_count}")
    print(f"Number of non-empty rows written to Excel Current Form List: {len(entries)}")
    print("✅ Current Form List Counts match." if rtf_count == len(entries) else "❌ Current Form List Counts do not match.")

def formlisttxt2xlsx_prev(self):
    txtfile = self.entry1.get()      
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    base_dir = os.path.join(initial_dir, txtfile, "FOD")
    report_dir = os.path.join(base_dir, 'Report')
    os.makedirs(report_dir, exist_ok=True)

    input_path = os.path.join(report_dir, "Previous_FORMLIST.txt")
    output_path = os.path.join(report_dir, "Previous_FORMLIST.xlsx")

    # Read the file content
    with open(input_path, "r", encoding="utf-8") as file:
        lines = file.readlines()

    entries = []
    rtf_count = 0

    # Extract relevant entries
    for line in lines:
        line = line.strip()
        if ".rtf" in line:
            rtf_count += 1
        if line and '|' in line and not line.startswith("File Name:"):
            parts = line.split('|')
            if len(parts) >= 3:
                file_name = parts[0].strip()
                form_number = parts[1].strip()
                form_title = parts[2].strip()
                if file_name and form_number and form_title:
                    entries.append([file_name, form_number, form_title])

    # Write to Excel (.xlsx)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Form List"
    ws.append(["File Name", "Form Number", "Form Title"])
    for entry in entries:
        ws.append(entry)
    wb.save(output_path)

    # Compare counts
    print(f"Number of '.rtf' entries in previous form list txt file: {rtf_count}")
    print(f"Number of non-empty rows written to Excel previous form list: {len(entries)}")
    print("✅ Previous Form List Counts match." if rtf_count == len(entries) else "❌ Previous Form List Counts do not match.")


# if __name__ == "__main__":
#     input_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\test\formlist\Current_FORMLIST.txt"
#     output_path = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive\test\formlist\form_listtxt2csv.csv"
#     formlisttxt2xlsx_prev(input_path, output_path)
#     # formlisttxt2xlsx(input_path, output_path) 