import os

def generate_report(self=None):
    entry1_text = self.entry1() if callable(self.entry1) else self.entry1.get()    
    # initial_dir = r"C:\Users\LABRADBM\Downloads\Local\YB\Python\FOD\I drive"
    initial_dir = r"\\fabwebd5.net\neptune\DataConversion\Prod\Secondary" 
    # base_dir = os.path.join(initial_dir, folder_entry)
    base_dir = os.path.join(initial_dir, entry1_text, "FOD")
    report_directory = os.path.join(base_dir, 'Report')    

    # Define mapping of month numbers to month names
    month_names = {
        1: "January", 2: "February", 3: "March", 4: "April",
        5: "May", 6: "June", 7: "July", 8: "August",
        9: "September", 10: "October", 11: "November", 12: "December"
    }
    # Create a reverse mapping for textual month names (case-insensitive) to numbers.
    month_name_to_number = {name.lower(): num for num, name in month_names.items()}

    # Retrieve month and year from widget attributes if available, otherwise fallback to interactive input.
    if self is not None and hasattr(self, "entry3_var") and hasattr(self, "entry4"):
        try:
            month_input = self.entry3_var.get().strip()
            # Check if the month input is numeric, otherwise assume it is textual
            if month_input.isdigit():
                month = int(month_input)
            else:
                # Convert text to lower case and map to its month number.
                month = month_name_to_number.get(month_input.lower())
                if month is None:
                    raise ValueError(f"Invalid month text: {month_input}")
            
            year = int(self.entry4.get().strip())
            
            if not (1 <= month <= 12):
                raise ValueError("Month must be between 1 and 12.")
            if year <= 0:
                raise ValueError("Year must be a positive number.")
        except Exception as e:
            print(f"Error retrieving month or year from entries: {e}")
            return
    else:
        # Fallback to interactive input if widget attributes are not provided.
        while True:
            try:
                month_input = input("Enter the month (1-12 or name, e.g., July): ").strip()
                if month_input.isdigit():
                    month = int(month_input)
                else:
                    month = month_name_to_number.get(month_input.lower())
                    if month is None:
                        raise ValueError("Invalid month name.")
                if 1 <= month <= 12:
                    break
                else:
                    print("Invalid month. Please enter a valid month number or name.")
            except ValueError as ve:
                print(f"Error: {ve}")

        while True:
            try:
                year = int(input("Enter the year (e.g., 2023): "))
                if year > 0:
                    break
                else:
                    print("Invalid year. Please enter a positive numerical value for the year.")
            except ValueError:
                print("Invalid input. Please enter a numerical value for the year.")

    # Get the month name based on the provided month number
    month_name = month_names.get(month, "Unknown Month")

    # Filename remains constant for release notes
    filename = "ReleaseNotes.txt"

    # Build the report content with month text and entry1 details if needed.
    report_content = f"""
Release Notes:
Changes to Forms on Download Release {self.entry2.get()}, {month_name} {year}
"""

    try:
        # Create the full path using the initial directory
        full_path = os.path.join(report_directory, filename)

        # Write the report file. Overwrites any existing file with the same name.
        with open(full_path, 'w', encoding="utf-8") as f:
            f.write(report_content.strip())  # Removes leading/trailing whitespace

        print(f"Report '{filename}' created successfully in '{report_directory}'")
    except IOError as e:
        print(f"Error releasenotes writing file '{filename}': {e}")
