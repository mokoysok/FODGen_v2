import os
import sys
import aspose.words as aw
from datetime import date
import tkinter.messagebox as messagebox
import sys

def compare_rtf_files(file_path1, file_path2, compared_output, summary_output):
    try:
        doc1 = aw.Document(file_path1)
        doc2 = aw.Document(file_path2)
        
        doc1.compare(doc2, "ComparisonUser", date.today())
        
        if doc1.revisions.count > 0:
            doc1.save(compared_output)
            print(f"Compared document saved to: {compared_output}")
            
            compared_doc = aw.Document(compared_output)
            revision_texts = set()
            output_lines = []
            
            for rev in compared_doc.revisions:
                if rev.revision_type in (aw.RevisionType.INSERTION, aw.RevisionType.DELETION):
                    parent_text = rev.parent_node.to_string(aw.SaveFormat.TEXT).strip()
                    normalized_text = ' '.join(parent_text.split())
                    if not normalized_text:
                        continue                    
                    if rev.revision_type == aw.RevisionType.INSERTION:
                        summary = f"Insertion: {normalized_text}"
                    else:
                        summary = f"Deletion: {normalized_text}"
                    
                    if summary not in revision_texts:
                        revision_texts.add(summary)
                        output_lines.append(summary)
            
            with open(summary_output, 'w', encoding='utf-8') as out_txt:
                out_txt.write("\n".join(output_lines))
            print(f"Revisions summary saved to: {summary_output}")
        else:
            print(f"No differences found between:\n  {file_path1}\n  {file_path2}")
    except Exception as e:
        print(f"Error comparing files '{file_path1}' and '{file_path2}': {e}")

def get_rtf_files(folder):
    files = {}
    for item in os.listdir(folder):
        if item.lower().endswith(".rtf"):
            files[item] = os.path.join(folder, item)
    return files

def comparertf(self):
    entry1_value = self.entry1.get()
    # print(entry1_value)
    entry2_value = self.entry5.get().strip().replace('.zip', '')
    # print(entry2_value)
    # Define the initial directory, update this path as needed
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    base_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary"
    initial_dir = os.path.join(base_dir, entry1_value, "FOD")
    FOLDER1 = os.path.join(initial_dir, entry2_value, "MODIFIED")
    # FOLDER2 = os.path.join(initial_dir, entry1_value)
    FOLDER2 = os.path.join(initial_dir, "MODIFIED")

    # REPORT_DIR = os.path.join(initial_dir, entry1_value, "Report")
    REPORT_DIR = os.path.join(initial_dir, "Report")
    if not os.path.exists(REPORT_DIR):
        os.makedirs(REPORT_DIR)

    folder1_files = get_rtf_files(FOLDER1)
    folder2_files = get_rtf_files(FOLDER2)
    
    common_files = set(folder1_files.keys()) & set(folder2_files.keys())
    
    if not common_files:
        response = messagebox.askyesno(
            "No Matching RTF",
            "No matching RTF files found in the two folders.\nDo you want to continue?"
        )
        if not response:
            sys.exit(0)  # Exit the program
        else:
            return  # Skip to the next iteration
    
    for filename in common_files:
        file1_path = folder1_files[filename]
        file2_path = folder2_files[filename]
        
        compared_filename = f"{os.path.splitext(filename)[0]}.rtf"
        summary_filename = f"{os.path.splitext(filename)[0]}.txt"
        
        compared_output = os.path.join(REPORT_DIR, compared_filename)
        summary_output = os.path.join(REPORT_DIR, summary_filename)
        
        print(f"Comparing file '{filename}'...")
        compare_rtf_files(file1_path, file2_path, compared_output, summary_output)
    
    folder1_only = set(folder1_files.keys()) - common_files
    folder2_only = set(folder2_files.keys()) - common_files
    
    if folder1_only:
        print("Files found only in Folder1:")
        for f in folder1_only:
            print("  " + f)
        folder1_report = os.path.join(REPORT_DIR, "DeletedForms.txt")
        with open(folder1_report, 'w', encoding='utf-8') as f1_out:
            f1_out.write(f"Files found only in Folder1: {FOLDER1}\n")
            for f in folder1_only:
                f1_out.write(f + "\n")

    if folder2_only:
        print("Files found only in Folder2:")
        for f in folder2_only:
            print("  " + f)
        folder2_report = os.path.join(REPORT_DIR, "AddedForms.txt")
        with open(folder2_report, 'w', encoding='utf-8') as f2_out:
            f2_out.write(f"Files found only in Folder2: {FOLDER2}\n")
            for f in folder2_only:
                f2_out.write(f + "\n")
    
    print("\nAll comparisons and reports have been generated.")

