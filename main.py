import tkinter.ttk as ttk
import tkinter as tk
from tkinter import messagebox, PhotoImage, Label, filedialog
from PIL import Image, ImageTk
import pandas as pd
from openpyxl.utils import get_column_letter
from datetime import datetime
import os
from pathlib import Path
import sys
import time
# --- Added for Threading ---
import threading
import queue 
# ---------------------------

# --- Your Module Imports ---
from functions import (
    CollectedForms,
    ExtractTitle,
    clean_up,
    compare_formlistnew,
    copyFinalguide,
    copy_ft_output,
    copyguide,
    deleteMODIFIED,
    extract_zipfile,
    formlisttxt2xlsxcitrix,
    listtxtnew,
    move_rtf,
    releasenotes,
    rename_rtf_frm_xlsx,
    revisionsforms,
    rtf2txt,
    userguideupdate,
    zip_rtf,
)
# ---------------------------

# Resolve bundled visual assets relative to the running script or executable.
APP_DIR = (
    Path(sys.executable).resolve().parent
    if getattr(sys, "frozen", False)
    else Path(__file__).resolve().parent
)
ICON_PATH = APP_DIR / "logo.ico"
LOGO_PATH = APP_DIR / "FOD.png"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("FOD User Guide Generator®")
        self.geometry("500x480")
        self.resizable(False, False)  
        self.configure(bg="#2a0b05", highlightbackground="black", borderwidth=1, relief="solid", highlightthickness=1, highlightcolor="white")
        self.iconbitmap(str(ICON_PATH))
        
        # --- Initialize Threading Attributes ---
        self.progress_queue = None
        self.worker_thread = None
        self.zip_queue = None      # Queue for the zip process
        self.zip_thread = None      # Thread for the zip process
        self.popup_window = None    # To reference the current pop-up window globally
        self.popup_process_label = None # To update status in the pop-up
        # --------------------------------------

        # Create a menu bar
        main_menu = tk.Menu(self)
        self.config(menu=main_menu)
        main_menu.add_cascade(label=" " * 146, state="disabled")  # Spacer
        def show_about():
            messagebox.showinfo(
            "About",
            "This tool will generate FOD User Guides for New/Deleted/Revised Forms.\nVersion: 1.0\nDeveloped by: BRi\nFor issues and concerns, please contact: brian.labrador@reedelsevier.com"
            )
        def show_guide():
            messagebox.showinfo(
            "Guide",
            "1. Enter 5-digit Pub Number.\n2. Enter Release Number.\n3. Enter/Select Revenue Month and Year.\n4. Browse and select last release zip file. \n   (Must be in the pubnum/FOD folder)\n5. Browse and select userguide reference. \n   (Must be in the pubnum/FOD folder)\n6. Enter path for [FT_OUTPUT_PATH] from job summary email. \n7. Click Generate to proceed.\n8.Click ZipFile to create Zip file of the current release.\n9. Click Exit to close the application."
            )

        help_menu = tk.Menu(main_menu, tearoff=0)
        help_menu.add_command(label="About", command=show_about)
        help_menu.add_command(label="Guide", command=show_guide)
        main_menu.add_cascade(label="Help", menu=help_menu)
        # Main frame  
        self.main_frame = tk.Frame(self, bg="#223556", width=150 , height=50)
        self.main_frame.pack(side="right", expand=True, fill="both", padx=10, pady=10)
        self.main_frame1 = tk.Frame(self.main_frame, bg="#223556", width=300, height=50)
        self.main_frame1.pack(side="top", expand=True, fill="both", padx=5, pady=5)
        self.image_file = str(LOGO_PATH)
        self.image = Image.open(self.image_file)
        self.image = self.image.resize((80, 80))  # Resize to desired size
        self.photo = ImageTk.PhotoImage(self.image)
        self.image_label = Label(self.main_frame1, image=self.photo, bg="#223556")
        self.image_label.pack(side="left", padx=10)
        self.main_label = Label(self.main_frame1, text="FOD User Guide Generator", font=("Corporate", 18), bg="#223556", fg="white")
        self.main_label.pack(side="left", padx=10, pady=10)
        self.sub_frame1 = tk.Frame(self.main_frame, bg="#223556")
        self.sub_frame1.pack(side="top", expand=False, fill="both", padx=5, pady=5)
        self.sub_frame2 = tk.Frame(self.main_frame, bg="#223556")
        self.sub_frame2.pack(side="top", expand=False, fill="both", padx=3, pady=3)        

        self.entry_label1 = tk.Label(self.sub_frame1, text="Input Pub Information!", font=("Arial", 9), bg="#223556", fg="white")
        self.entry_label1.pack(pady=(5, 0), side="top", anchor="center")

        entry_fields_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields_frame.pack(pady=2, side="top", anchor="center", expand=False)
        entry_fields1_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields1_frame.pack(pady=2, side="top", anchor="center", expand=False)
        entry_fields2_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields2_frame.pack(pady=2, side="top", anchor="center", expand=False)        
        entry_fields3_frame = tk.Frame(self.sub_frame1, bg="#223556")
        entry_fields3_frame.pack(pady=2, side="top", anchor="center", expand=False)          

        entry1_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry1_frame.grid(row=0, column=0, padx=10, pady=2, sticky="w")
        self.entry1_label = tk.Label(entry1_frame, text="Pub No. (5-digits):", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry1_label.pack(side="left", padx=(5, 0))
        self.entry1 = tk.Entry(entry1_frame, width=15, font=("Arial", 8))
        self.entry1.pack(side="left", fill="x", expand=False)

        entry2_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry2_frame.grid(row=0, column=1, padx=10, pady=2, sticky="w")
        self.entry2_label = tk.Label(entry2_frame, text="Rel Number:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry2_label.pack(side="left", padx=(5, 0))
        self.entry2 = tk.Entry(entry2_frame, width=15, font=("Arial", 8))
        self.entry2.pack(side="left", fill="x", expand=False)        

        entry3_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry3_frame.grid(row=1, column=0, padx=10, pady=2, sticky="w")
        self.entry3_label = tk.Label(entry3_frame, text="Rev Month:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry3_label.pack(side="left", padx=(5, 0))
        # Dropdown for months
        months = ["January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"]
        self.entry3_var = tk.StringVar()
        self.entry3_var.set(months[0])  # Default to current month
        current_month_index = datetime.now().month - 1  # 0-based index
        self.entry3 = ttk.Combobox(entry3_frame, textvariable=self.entry3_var, values=months, state="readonly", width=12, font=("Arial", 8))
        self.entry3.current(current_month_index)
        self.entry3_var.set(months[current_month_index])
        self.entry3.pack(side="right", fill="x", expand=False)

        entry4_frame = tk.Frame(entry_fields_frame, bg="#F7F7F8")
        entry4_frame.grid(row=1, column=1, padx=10, pady=2, sticky="w")
        self.entry4_label = tk.Label(entry4_frame, text="Year:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry4_label.pack(side="left", padx=(5, 0))
        self.entry4 = tk.Entry(entry4_frame, width=15, font=("Arial", 8))
        self.entry4.pack(side="left", fill="x", expand=True)
        self.entry4.insert(0, str(datetime.now().year))

        entry5_frame = tk.Frame(entry_fields2_frame, bg="#F7F7F8")
        entry5_frame.pack(pady=2, side="top", anchor="center", expand=False)
        self.entry5_label = tk.Label(entry5_frame, text="Previous Release ZipFile:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=20)
        self.entry5_label.pack(side="left", padx=(5, 0))
        self.entry5 = tk.Entry(entry5_frame, width=37, font=("Arial", 8))
        self.entry5.pack(side="left", padx=5)
        def browse_zip_file():
            #C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive
            initial_dir = os.path.join(
                r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary",
                self.entry1.get().strip(),"FOD"
            )
            if not os.path.isdir(initial_dir):
                initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary"
            file_path = filedialog.askopenfilename(
                initialdir=initial_dir,
                title="Select Zip File",
                filetypes=[("Zip Files", "*.zip"), ("All Files", "*.*")]
            )
            if file_path:
                self.entry5.delete(0, tk.END)
                self.entry5.insert(0, file_path)
        self.entry5_button = tk.Button(entry5_frame, text="Browse", font=("Arial", 8), command=browse_zip_file, border=1, relief="solid", bg="#F7F7F8", fg="black")
        self.entry5_button.pack(side="right", padx=1)
        
        entry6_frame = tk.Frame(entry_fields2_frame, bg="#F7F7F8")
        entry6_frame.pack(pady=2, side="top", anchor="center", expand=False)
        self.entry6_label = tk.Label(entry6_frame, text="Current User Guide:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=20)
        self.entry6_label.pack(side="left", padx=(5, 0))
        self.entry6 = tk.Entry(entry6_frame, width=37, font=("Arial", 8))
        self.entry6.pack(side="left", padx=5)
        def browse_userguide():
            #C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive
            initial_dir = os.path.join(
                r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary",
                self.entry1.get().strip(), "FOD"
            )  
            
            if not os.path.isdir(initial_dir):
                initial_dir = r"\\legal.regn.net\FAB\Neptune\Prod\DataConversion\Prod\Secondary"
            file_path = filedialog.askopenfilename(
                initialdir=initial_dir,
                title="Select User Guide",
                filetypes=[("RTF Files", "*.rtf"), ("All Files", "*.*")]
            )
            if file_path:
                self.entry6.delete(0, tk.END)
                self.entry6.insert(0, file_path)
        self.entry6_button = tk.Button(entry6_frame, text="Browse", font=("Arial", 8), command=browse_userguide, border=1, relief="solid", bg="#F7F7F8", fg="black")
        self.entry6_button.pack(side="right", padx=1)

        entry7_frame = tk.Frame(entry_fields3_frame, bg="#F7F7F8")
        entry7_frame.pack(pady=2, side="top", anchor="center", expand=False)
        self.entry7_label = tk.Label(entry7_frame, text="FT_OUTPUT_PATH:", font=("Arial", 8), bg="#F7F7F8", fg="black", width=15)
        self.entry7_label.pack(side="left", padx=(5, 0))
        self.entry7 = tk.Entry(entry7_frame, width=50, font=("Arial", 8))
        self.entry7.pack(side="left", padx=5)  

        self.submit_button = tk.Button(self.sub_frame2, text="Generate", font=("Gagalin", 10), command=self.check_path, width=17, height=2, border=2, relief="solid", bg="#F7F7F8", fg="black")
        self.submit_button.pack(pady=15, padx=10, side="left", anchor="w")
        self.Revoke_button = tk.Button(self.sub_frame2, text="Zip File", command=self.check_pathrevoke, width=17, height=2, border=2, relief="solid", bg="#F7F7F8", fg="black")
        self.Revoke_button.pack(padx=15, side="left", anchor="w")
        # self.rerun_button = tk.Button(self.sub_frame2, text="Rerun", command=self.refresh_run, width=13, height=2)
        # self.rerun_button.pack(padx=10, side="left", anchor="w")
        self.exit_app_button = tk.Button(self.sub_frame2, text="Exit", command=self.exit_app, width=17, height=2, border=2, relief="solid", bg="#F7F7F8", fg="black")
        self.exit_app_button.pack(padx=15, side="left", anchor="w")

        self.sub_frame4 = tk.Frame(self.main_frame, bg="gray", width=150, height=100)
        self.sub_frame4.pack(side="top", expand=False, fill="both", padx=5, pady=5) 
        self.sub_frame3 = tk.Frame(self.main_frame, bg="#223556")
        self.sub_frame3.pack(side="top", expand=True, fill="both", padx=5, pady=5)           
        
        self.percent_label = Label(self.sub_frame3, text="0%", font=("Arial", 10), bg="#223556", fg="white", height=1) # Font size corrected
        self.percent_label.pack(pady=1, side="top", anchor="n")          
        self.progress_bar = tk.Scale(self.sub_frame4, orient="horizontal", length=500, sliderlength=0, # sliderlength changed to 0
                                     showvalue=0, from_=0, to=100, 
                                     bg="#ECEEF3", highlightbackground="#223556")
        self.progress_bar.set(0)  
        self.progress_bar.pack(side="top", anchor="w", fill='x', expand=True)

        self.start_time = 0

    def exit_app(self):
        self.destroy()

    def update_timer(self, timer_label, popup):
        # This function is not used in the threading approach but kept for reference
        elapsed_time = time.time() - self.start_time
        minutes, seconds = divmod(int(elapsed_time), 60)
        timer_label.configure(text=f"Time Elapsed: {minutes:02}:{seconds:02}")
        if popup.winfo_exists():
            popup.after(1000, self.update_timer, timer_label, popup)         
            
    def get_elapsed_time(self):
        elapsed = int(time.time() - self.start_time)
        minutes, seconds = divmod(elapsed, 60)
        return f"{minutes:02}:{seconds:02}"

    # ----------------------------------------------------------------------
    ## GENERATE PROCESS (check_path and run_all_functions)
    # ----------------------------------------------------------------------

    def check_path(self):
        """Starts the main 'Generate' process, using a background thread."""
        
        pub_number = self.entry1.get().strip()
        rel_number = self.entry2.get().strip()
        rev_month = self.entry3_var.get().strip()
        year = self.entry4.get().strip()
        prev_zip = self.entry5.get().strip()
        userguide = self.entry6.get().strip()
        ft_output = self.entry7.get().strip()

        if not all([pub_number, rel_number, rev_month, year, prev_zip, userguide, ft_output]):
            tk.messagebox.showerror("Error", "All fields must be filled in.")
            return

        if tk.messagebox.askyesno("Confirm", "All fields are filled in. Start processing?"):
            
            self.progress_bar.set(0)
            self.percent_label.configure(text="0%")
            
            self._set_button_state("disabled", "Processing...")

            # 1. Setup pop-up window
            self.popup_window = self._create_popup("Processing...")
            
            # 2. Setup Queue
            self.progress_queue = queue.Queue()
            
            # 3. Setup and start Thread
            self.worker_thread = threading.Thread(target=self.run_all_functions)
            self.worker_thread.daemon = True
            self.worker_thread.start()
            
            # 4. Start main thread periodic check
            self.after(100, self.check_progress_queue)

        else:
            tk.messagebox.showerror("Error", "Process cancelled by user.")

    def run_all_functions(self):
        """Executes all MAIN processing steps in the background thread."""
        
        functions_to_run = [
            (move_rtf.move_rtf_files, "Moving RTF files..."),
            (copy_ft_output.copy_ftoutput_rtf_files, "Copying FT output RTF files..."),
            (extract_zipfile.extract_zip_action, "Extracting zip file..."),
            (rtf2txt.convert_rtf_to_txt_current, "Converting current RTF to TXT..."),
            (rtf2txt.convert_rtf_to_txt_prev, "Converting previous RTF to TXT..."),
            (formlisttxt2xlsxcitrix.formlisttxt2xlsx_prev, "Creating Excel from previous form list..."),
            (formlisttxt2xlsxcitrix.formlisttxt2xlsx_current, "Creating Excel from current form list..."),
            (rename_rtf_frm_xlsx.rename_rtf_from_excel1, "Renaming RTF file 1..."),
            (rename_rtf_frm_xlsx.rename_rtf_from_excel2, "Renaming RTF file 2..."),
            (revisionsforms.comparertf, "Comparing RTF revisions..."),
            (copyguide.copy_file, "Copying guide file..."),
            (listtxtnew.listrevision_txt_files, "Listing revision TXT files..."),
            (listtxtnew.revised2xlsx, "Creating Excel from revised list..."),
            (listtxtnew.match_form_titles, "Matching form titles..."),
            (compare_formlistnew.compare_formlists, "Comparing form lists..."),
            (compare_formlistnew.match_form_titles_add, "Matching added form titles..."),
            (compare_formlistnew.match_form_titles_deleted, "Matching deleted form titles..."),
            (CollectedForms.process_forms, "Processing collected forms..."),
            (ExtractTitle.extract_forms_title_from_rtf, "Extracting form titles from RTF..."),
            (ExtractTitle.extract_forms_sample_from_rtf, "Extracting form samples from RTF..."),
            (releasenotes.generate_report, "Generating release notes report..."),
            (deleteMODIFIED.delete_modified_userguide_rtf, "Deleting MODIFIED user guide RTF..."),
            (userguideupdate.userguide_update, "Updating user guide...")
        ]
        
        total_steps = len(functions_to_run)
        
        try:
            start_time = time.time()
            
            for i, (func, description) in enumerate(functions_to_run):
                self.progress_queue.put(("STATUS", description))
                func(self)
                percent_complete = int(((i + 1) / total_steps) * 99)
                self.progress_queue.put(percent_complete)
                
            end_time = time.time()
            elapsed_time = end_time - start_time
            
            self.progress_queue.put(100) 
            self.progress_queue.put(("COMPLETE", elapsed_time))
            
        except Exception as e:
            self.progress_queue.put(("ERROR", str(e)))
            
    # ----------------------------------------------------------------------
    ## ZIP PROCESS (check_pathrevoke and run_zip_functions)
    # ----------------------------------------------------------------------
    
    def check_pathrevoke(self):
        """Starts the 'Zip File' process, using a background thread."""
        
        pub_number = self.entry1.get().strip()
        rel_number = self.entry2.get().strip()

        if not all([pub_number, rel_number]):
            messagebox.showerror("Error", "Please enter Pub Number and Rel Number.")
            return

        if messagebox.askyesno("Confirm", "Zip rtf files?"):

            self.progress_bar.set(0)
            self.percent_label.configure(text="0%")

            self._set_button_state("disabled", "Processing...")
            
            # 1. Setup pop-up window
            self.popup_window = self._create_popup("Zipping Files...")
            
            # 2. Setup Queue 
            self.zip_queue = queue.Queue()
            
            # 3. Setup and start Thread
            self.zip_thread = threading.Thread(target=self.run_zip_functions)
            self.zip_thread.daemon = True
            self.zip_thread.start()
            
            # 4. Start main thread periodic check for the zip process
            self.after(100, self.check_progress_queue_zip)
        else:
            messagebox.showerror("Error", "Process cancelled by user.")

    def run_zip_functions(self):
        """Executes the ZIP processing steps in the background thread."""
        
        functions_to_run = [
            (copyFinalguide.copy_finalfile, "Updating User guide file..."),
            (zip_rtf.zip_rtf_files, "Zipping RTF files..."),
            (clean_up.clean_directories, "Cleaning up directories...")
        ]
        
        total_steps = len(functions_to_run)
        
        try:
            start_time = time.time()
            
            for i, (func, description) in enumerate(functions_to_run):
                self.zip_queue.put(("STATUS", description))
                func(self)
                percent_complete = int(((i + 1) / total_steps) * 99)
                self.zip_queue.put(percent_complete)
                
            end_time = time.time()
            elapsed_time = end_time - start_time
            
            self.zip_queue.put(100)
            self.zip_queue.put(("COMPLETE", elapsed_time))
            
        except Exception as e:
            self.zip_queue.put(("ERROR", str(e)))

    # ----------------------------------------------------------------------
    ## SHARED UTILITY METHODS (GUI UPDATE & POPUP)
    # ----------------------------------------------------------------------

    def _create_popup(self, title):
        """Creates and centers a non-blocking processing pop-up window."""
        popup = tk.Toplevel(self)
        popup.title(title)
        # Set the window icon
        icon_path = str(ICON_PATH)
        try:
            popup.iconbitmap(icon_path)
        except Exception as e:
            print(f"Failed to set icon: {e}")        

        popup.transient(self)
        popup.grab_set()

        # Center the pop-up:
        main_window_x = self.winfo_x()
        main_window_y = self.winfo_y()
        main_window_width = self.winfo_width()
        main_window_height = self.winfo_height()
        popup_width = 300
        popup_height = 100
        popup_x = main_window_x + (main_window_width // 2) - (popup_width // 2)
        popup_y = main_window_y + (main_window_height // 2) - (popup_height // 2)
        popup.geometry(f"{popup_width}x{popup_height}+{popup_x}+{popup_y}")
        popup.resizable(False, False)

        # Create a label that shows the current processing step:
        process_label = tk.Label(popup, text="Starting process...", font=("Arial", 10))
        process_label.pack(pady=10)
        self.popup_process_label = process_label  # Save for easy status update
        
        return popup

    def _set_button_state(self, state, submit_text):
        """A utility method to quickly set button states."""
        self.submit_button.configure(state=state, text=submit_text)
        self.Revoke_button.configure(state=state, text="Zip File")
        self.exit_app_button.configure(state=state, text="Exit")


    def _handle_queue_update(self, current_queue, is_zip_process=False):
        """Centralized logic to read from a queue and update the GUI."""
        
        try:
            while True:
                item = current_queue.get_nowait()
                
                if isinstance(item, int):
                    # Percentage update
                    percent = item
                    self.progress_bar.set(percent)
                    self.percent_label.configure(text=f"{percent}%")

                elif isinstance(item, tuple) and item[0] == "STATUS":
                    # Status message update
                    description = item[1]
                    self.percent_label.configure(text=f"{self.progress_bar.get()}% - {description}")
                    if self.popup_process_label and self.popup_window and self.popup_window.winfo_exists():
                        self.popup_process_label.configure(text=description)

                elif isinstance(item, tuple) and item[0] == "COMPLETE":
                    # Final completion signal
                    elapsed_time = item[1]
                    
                    if self.popup_window and self.popup_window.winfo_exists():
                        self.popup_window.destroy()

                    tk.messagebox.showinfo("Complete", f"Process Complete!\nElapsed Time: {elapsed_time:.2f} seconds")
                    
                    # Final state for both processes
                    if is_zip_process:
                        self._set_button_state("normal", "Regenerate")
                    else:
                        # After a successful Generate, enable Zip File button
                        self.submit_button.configure(state="disabled", text="Done")
                        self.Revoke_button.configure(state="normal", text="Zip File")
                        self.exit_app_button.configure(state="normal", text="Exit")
                        
                    self.progress_bar.set(100)
                    self.percent_label.configure(text="100% Process Complete!")
                    return # Stop re-scheduling

                elif isinstance(item, tuple) and item[0] == "ERROR":
                    # Handle error signal
                    error_message = item[1]
                    
                    if self.popup_window and self.popup_window.winfo_exists():
                        self.popup_window.destroy()
                        
                    tk.messagebox.showerror("Error", f"An error occurred:\n{error_message}")
                    
                    # Reset all buttons on error
                    self._set_button_state("normal", "Generate")
                    self.progress_bar.set(0)
                    self.percent_label.configure(text="0%")
                    return # Stop re-scheduling
            
        except queue.Empty:
            # Normal state: no updates in the queue right now
            pass
        
        # If we reach here, the process is still running, so reschedule the check
        if is_zip_process:
            self.after(100, self.check_progress_queue_zip)
        else:
            self.after(100, self.check_progress_queue)


    # Main thread check for the 'Generate' process
    def check_progress_queue(self):
        """Checks the main progress queue (Generate process) and updates the GUI."""
        if self.progress_queue:
            self._handle_queue_update(self.progress_queue, is_zip_process=False)
        
    # Main thread check for the 'Zip File' process
    def check_progress_queue_zip(self):
        """Checks the zip progress queue (Zip File process) and updates the GUI."""
        if self.zip_queue:
            self._handle_queue_update(self.zip_queue, is_zip_process=True)


    # def refresh_run(self):
    #     self.submit_button.configure(state="normal", text="Submit")  
    #     # self.entry.delete(0, tk.END) # Assuming this was a placeholder for an entry not in the provided code
    #     messagebox.showinfo("Re-run", "You can do another Clean up.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
