import tkinter as tk #Tkinter library for building the GUI
from tkinter import filedialog, messagebox #file dialog and messagebox for interaction
import pandas as pd #Pandas library for handling Excel fles
import xlsxwriter #XlsxWriter library for writing Excel files with formatting
import math #module to allow for mathematical functions and constants
import re #Regular expression module used for pattern-based string manipulation
import csv #for reading from and writing to CSV
import numpy as np #numerical python
import unicodedata #module for the Unicode Character Database
from pathlib import Path #module for modern object-oriented way to handle filesystem paths

#Defines a class to compare files using graphical interface
class CompareFiles:
    def __init__(self, root):
        self.root = root #assigns Tkinter root window
        self.root.title("Audit Comparison Tool") #window title
        self.root.configure(bg="#2b2b2b") #set window background color to white

        self.wagerauditStaging_path = "" #path for Wager Staging Audit csv file
        self.wagerauditProduction_path = "" #path for Wager Production Audit csv file
        self.operatorsheet_path = "" #path for Op Wager Config Excel sheet file
        self.opgamelistStaging_report_path = "" #path for Op Staging GameList csv Report
        self.opgamelistProduction_report_path = "" #path for Op Production GameList csv Report
        self.agilereport_path = "" #path for the Agile PLM Excel Report
        self.create_widgets() #method to create UI components

        self.adjust_window(root) #set window size and center it
        self.root.geometry("800x600") #setting default and min size settings
        self.root.minsize(800, 600)

    def adjust_window(self, root):
        #Get the screen's full width/height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 800
        window_height = 600

        #Calculate position to center the window
        position_top = (screen_height - window_height) // 2
        position_left = (screen_width - window_width) // 2

        #Update the window's geometry to apply size and position
        root.geometry(f'{screen_width}x{window_height}+{position_left}+{position_top}')

    def create_widgets(self):
        #Main content frame for all buttons/labels
        content_frame = tk.Frame(self.root, bg="#2b2b2b", height=300)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)

        #Welcome display text and label
        welcome_text = "\nAudit Comparison Tool\n"
        self.welcome_label = tk.Label(content_frame, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='white', bg='#2b2b2b')
        self.welcome_label.pack(pady=10)

        #Container for left/right groups side by side
        group_container = tk.Frame(content_frame, bg="#2b2b2b")
        group_container.pack()

        #Left group (wager audit file/op wager config sheet)
        left_group = tk.LabelFrame(group_container, bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
        left_group.pack(side="left", padx=10)

        #Right group (op gamelist report/agile plm report)
        right_group = tk.LabelFrame(group_container, bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
        right_group.pack(side="right", padx=10)     

        #Button style dictionary for all buttons
        button_style = {
            "bg": "#6e6e6e",
            "fg": "white",
            "activebackground": "#505050",
            "activeforeground": "white",
            "borderwidth": 1,
            "highlightthickness": 0,
            "font": ("TkDefaultFont", 10, "bold")
        }
        
        #Label styles dictionary for all labels
        label_style = {
            "bg": "#2b2b2b",
            "fg": "#FF6F6F",
            "padx": 10,
            "pady": 2,
            "font": ("TkDefaultFont", 10)
        }

        #Wager Audit from Staging label and upload button
        self.wagerauditStaging_label = tk.Label(left_group, text="Select Wager Staging Audit csv file", **label_style)
        self.wagerauditStaging_label.pack(pady=(0, 5))
        self.upload_wagerauditStaging = tk.Button(left_group, text="Upload Wager Staging Audit csv file", command=self.upload_wagerauditStaging, **button_style)
        self.upload_wagerauditStaging.pack(pady=(0, 10))
        self.button_hover_effect(self.upload_wagerauditStaging)

        #Wager Audit from Production label and upload button
        self.wagerauditProduction_label = tk.Label(left_group, text="Select Wager Production Audit csv file", **label_style)
        self.wagerauditProduction_label.pack(pady=(10, 5))
        self.upload_wagerauditProduction = tk.Button(left_group, text="Upload Wager Production Audit csv file", command=self.upload_wagerauditProduction, **button_style)
        self.upload_wagerauditProduction.pack(pady=(0, 10))
        self.button_hover_effect(self.upload_wagerauditProduction)

        #Operator Wager Config Sheet label and upload button
        self.operatorsheet_label = tk.Label(left_group, text="Select Operator Wager Config Excel Sheet", **label_style)
        self.operatorsheet_label.pack(pady=(10, 5))
        self.upload_operatorsheet = tk.Button(left_group, text="Upload Operator Wager Config Excel Sheet", command=self.upload_operatorsheet, **button_style)
        self.upload_operatorsheet.pack(pady=(0, 10))
        self.button_hover_effect(self.upload_operatorsheet)

        #Operator Staging GameList Report label and upload button
        self.opgamelistStaging_report_label = tk.Label(right_group, text="Select Operator Staging GameList csv Report", **label_style)
        self.opgamelistStaging_report_label.pack(pady=(0, 5))
        self.upload_opgamelistStaging_report = tk.Button(right_group, text="Upload Operator Staging GameList csv Report", command=self.upload_opgamelistStaging_report, **button_style)
        self.upload_opgamelistStaging_report.pack(pady=(0, 10))
        self.button_hover_effect(self.upload_opgamelistStaging_report)

        #Operator Production GameList Report label and upload button
        self.opgamelistProduction_report_label = tk.Label(right_group, text="Select Operator Production GameList csv Report", **label_style)
        self.opgamelistProduction_report_label.pack(pady=(10, 5))
        self.upload_opgamelistProduction_report = tk.Button(right_group, text="Upload Operator Production GameList csv Report", command=self.upload_opgamelistProduction_report, **button_style)
        self.upload_opgamelistProduction_report.pack(pady=(0, 10))
        self.button_hover_effect(self.upload_opgamelistProduction_report)

        #Agile PLM Report label and upload button
        self.agilereport_label = tk.Label(right_group, text="Select Agile PLM Excel Report", **label_style)
        self.agilereport_label.pack(pady=(10, 5))
        self.upload_agilereport = tk.Button(right_group, text="Upload Agile PLM Excel Report", command=self.upload_agilereport, **button_style)
        self.upload_agilereport.pack(pady=(0, 10))
        self.button_hover_effect(self.upload_agilereport)

        #Frame for submit/clear buttons
        action_frame = tk.Frame(content_frame, bg="#2b2b2b")
        action_frame.pack(pady=20)

        #Submit button
        self.submit_button = tk.Button(action_frame, text="SUBMIT FILES", font=("TkDefaultFont", 12, "bold"), command=self.submit_files, state=tk.DISABLED, fg='white', bg="#FF6F6F", borderwidth=1)
        self.submit_button.pack(side="left", padx=10)
        self.button_hover_effect(self.submit_button)

        #Clear button
        self.clear_button = tk.Button(action_frame, text="CLEAR FILES", font=("TkDefaultFont", 12, "bold"), command=self.clear_button, fg='white', bg="#6e6e6e", borderwidth=1)
        self.clear_button.pack(side="left", padx=10)
        self.button_hover_effect(self.clear_button)

        #Center action_frame for submit/clear buttons
        action_frame.pack_configure(anchor="center")

    #Adds a hover effect to the buttons
    def button_hover_effect(self, button, hover_bg="#5a5a5a", normal_bg="#6e6e6e"):
        #Takes into consideration the submit buttons color effects of green/red
        def on_enter(e):
            if button.cget("state") == tk.NORMAL:
                if button.cget("bg") != "green":
                    button.config(bg=hover_bg)

        def on_leave(e):
            if button.cget("state") == tk.NORMAL:
                if button.cget("bg") == hover_bg:
                    button.config(bg=normal_bg)
            else:
                button.config(bg="#FF6F6F")

        #Bind hover effect to the button
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def enable_submit_button(self):
        #Enables the submit button if files are not empty and turns green. Displays red if no files are selected and remains disabled
        if all([self.wagerauditStaging_path, self.wagerauditProduction_path, self.operatorsheet_path, self.opgamelistStaging_report_path, self.opgamelistProduction_report_path, self.agilereport_path]):
            self.submit_button.config(state=tk.NORMAL, bg='green')
        else:
            self.submit_button.config(state=tk.DISABLED, bg='#FF6F6F')
        self.button_hover_effect(self.submit_button)

    def upload_wagerauditStaging(self):
        self.wagerauditStaging_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")]) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerauditStaging_path: #Checks if a file is selected
            self.wagerauditStaging_label.config(text=f"Wager Staging Audit csv file uploaded: \n{self.wagerauditStaging_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Wager Staging Audit csv file to proceed.") #Show warning if no wager audit file is selected
            self.wagerauditStaging_label.config(text="Select Wager Staging Audit csv file", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.wagerauditStaging_path = "" if not self.wagerauditStaging_path else self.wagerauditStaging_path
        self.enable_submit_button() #Enables submit button after selection

    def upload_wagerauditProduction(self):
        self.wagerauditProduction_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")]) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerauditProduction_path: #Checks if a file is selected
            self.wagerauditProduction_label.config(text=f"Wager Production Audit csv file uploaded: \n{self.wagerauditProduction_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Wager Production Audit csv file to proceed.") #Show warning if no wager audit file is selected
            self.wagerauditProduction_label.config(text="Select Wager Production Audit csv file", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.wagerauditProduction_path = "" if not self.wagerauditProduction_path else self.wagerauditProduction_path
        self.enable_submit_button() #Enables submit button after selection 

    def upload_operatorsheet(self):
        self.operatorsheet_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")]) #Allows user to upload excel file (this is the file type when file is downloaded)

        if self.operatorsheet_path: #Checks if a file is selected
            self.operatorsheet_label.config(text=f"Operator Wager Config Excel Sheet uploaded: \n{self.operatorsheet_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Operator Wager Config Excel Sheet to proceed.") #Show warning if no operator wager config sheet is selected
            self.operatorsheet_label.config(text="Select Operator Wager Config Excel Sheet", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.operatorsheet_path = "" if not self.operatorsheet_path else self.operatorsheet_path
        self.enable_submit_button() #Enables submit button after selection

    def upload_opgamelistStaging_report(self):
        self.opgamelistStaging_report_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")]) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opgamelistStaging_report_path: #Checks if a file is selected
            self.opgamelistStaging_report_label.config(text=f"Operator Staging GameList csv Report uploaded: \n{self.opgamelistStaging_report_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Operator Staging GameList csv Report to proceed.") #Show warning if no op game list file is selected 
            self.opgamelistStaging_report_label.config(text="Select Operator Staging GameList csv Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opgamelistStaging_report_path = "" if not self.opgamelistStaging_report_path else self.opgamelistStaging_report_path
        self.enable_submit_button() #Enables submit button after selection
    
    def upload_opgamelistProduction_report(self):
        self.opgamelistProduction_report_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")]) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opgamelistProduction_report_path: #Checks if a file is selected
            self.opgamelistProduction_report_label.config(text=f"Operator Production GameList csv Report uploaded: \n{self.opgamelistProduction_report_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Operator Production GameList csv Report to proceed.") #Show warning if no op game list file is selected 
            self.opgamelistProduction_report_label.config(text="Select Operator Production GameList csv Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opgamelistProduction_report_path = "" if not self.opgamelistProduction_report_path else self.opgamelistProduction_report_path

    def upload_agilereport(self):
        self.agilereport_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")]) #Allows user to upload excel file (this is the file type when file is downloaded from agile power bi)

        if self.agilereport_path: #Checks if a file is selected
            self.agilereport_label.config(text=f"Agile PLM Excel Report uploaded: \n{self.agilereport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Agile PLM Excel Report to proceed.") #Show warning if no Agile PLM Report is selected
            self.agilereport_label.config(text="Select Agile PLM Excel Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.agilereport_path = "" if not self.agilereport_path else self.agilereport_path
        self.enable_submit_button() #Enables submit button after selection

    def submit_files(self):
        #Checks if all files are uploaded
        if not all([self.wagerauditStaging_path, self.wagerauditProduction_path, self.operatorsheet_path, self.opgamelistStaging_report_path, self.opgamelistProduction_report_path, self.agilereport_path]):
            messagebox.showwarning("Incomplete files!", "Please upload all required files before submitting.") #Show warning if not all files were uploaded
            return
        
        root = tk.Tk() #Creates a tkinter root window
        root.withdraw() #Hides the root window (so it doesn't pop up)

        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #file types filter
            title="File Save Location" #dialog title
        )

        if not file_path:
            messagebox.showinfo("Missing File Path!", "Select a file path to save the Final Audit Results and try again.") #Show cancelled message if no save file path was selected
            self.enable_submit_button() #Enables submit button
            return

        #Message box to confirm user selected files for submission and allows user to hit cancel if needed to re-upload files
        if messagebox.askyesno("Confirm", "Are you sure you want to submit the selected files for comparison?"):
            try:
                result = self.compare_files(file_path) #Call the function to compare files and save
                if result:
                    messagebox.showinfo("Audit Results Saved!", f"Audit Results file successfully saved at: {file_path}.") #Success message and show user save location
                else:
                    messagebox.showerror("Error!", "Failed to save file. Please check the correct file formats were submitted and try again.") #Show failure message if results fail
            except Exception as e:
                messagebox.showerror("Error!", f"Error occurred during export: {str(e)}") #Show error if there's an exception while saving files
        else:
            messagebox.showinfo("Cancelled!", "File submission cancelled. Please upload all required files to submit and try again.") #Display cancel message if user hits cancel

        self.enable_submit_button() #Resets submit button to it's default state after handling success, cancellation, or missing file path

    def clear_button(self):
        #Clear all file paths
        self.wagerauditStaging_path = ""
        self.wagerauditProduction_path = ""
        self.operatorsheet_path = ""
        self.opgamelistStaging_report_path = ""
        self.opgamelistProduction_report_path = ""
        self.agilereport_path = ""

        #Clear all labels and display red text
        self.wagerauditStaging_label.config(text="Select Wager Staging Audit csv file", fg="#FF6F6F")
        self.wagerauditProduction_label.config(text="Select Wager Production Audit csv file", fg="#FF6F6F")
        self.operatorsheet_label.config(text="Select Operator Wager Config Excel Sheet", fg="#FF6F6F")
        self.opgamelistStaging_report_label.config(text="Select Operator Staging GameList csv Report", fg="#FF6F6F")
        self.opgamelistProduction_report_label.config(text="Select Operator Production GameList csv Report", fg="#FF6F6F")
        self.agilereport_label.config(text="Select Agile PLM Excel Report", fg="#FF6F6F")

        self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F") #Disable the submit button and turn red

    def normalize_name(self, name):
        #Standardize game name column; convert to lowercase, removes all spaces, removes apostrophes
        if isinstance(name, str):
            name = unicodedata.normalize('NFKD', name) #Normalize any smart quotes or accents (Ex: Jack O'Lantern Jackpots)
            name = name.strip().replace(' ', '') #Remove spaces
            name = re.sub(r"[’']", '', name) #Remove straight and curly apostrophes using regex 
            return name.lower() #Convert to lowercase
        return name
    
    def normalize_value(self, val):
        #Standardize values to handle percentages, currencies, and NaN values
        if pd.isna(val) or val == '' or val == ' ': #Return empty string for NaN, empty string, or whitespace
            return ''
        val = str(val).strip()
        
        #Handles converting percentages to decimals (ex: 90% -> 0.9)
        if isinstance(val, str) and '%' in val:
            try:
                decimal_val = float(val.replace('%', '').strip()) / 100 #Convert to decimal
                return str(math.ceil(decimal_val * 100) / 100) #Rounds up to the next decimal place (ex: 0.9595 to 0.96); rounding can be removed if op wager sheets have exact RTPs; will highlight red if not exact
            except ValueError:
                return '' #if conversion fails, return empty string
        
        #Handles multiple values separated by commas or space separated values (ex: $0.01, $0.05, $0.10, etc.)
        if any(sym in val for sym in ('$', '€', '£')): #Can add more currencies as needed
            currency_values = re.findall(r'[\$€£]?\d[\d,]*\.?\d*', val) #Regex to detect multiple values vs single values
            
            if len(currency_values) > 1: #Checks for more than one currency value
                parts = [v.strip() for v in val.split(',')] #Split by commas and strip whitespace from each individual value
                normalized_values = [self.normalize_currency_values(p) for p in parts] #Normalize each stripped currency value using def normalize_currency_values methood
                return ','.join(normalized_values) #Join normalized values back into a single comma-separated string
            else:
                return self.normalize_currency_values(val) #Single value - normalize it directly
        
        val = val.replace(' ', '') #Remove all spaces
        return self.clean_number_string(val) #Clean string using def clean_number_string method
    
    def clean_number_string(self, val):
        #Handles values without currency symbols such as default lines & bet multipliers
        try:
            num = float(val) #Convert to float
            if num.is_integer(): #Checks if float is an integer
                return str(int(num)) #If integer, return as a string
            else:
                return str(num) #If not an integer, return float as a string
        except ValueError:
            return str(val).strip() #If conversion fails (val is not a number), return original value as a stripped string

    def normalize_currency_values (self, val):
    #Helper method to handle currency symbols (ex: $€£) and commas; can expand currencies as needed
        try:
            val = val.replace('$', '').replace('€', '').replace('£', '').replace(',', '').strip() #Remove the currency symbols/commas
            num = float(val) #Convert to float
            if num.is_integer(): #Checks if float is an integer
                return str(int(num)) #If integer, convert to an integer then to a string (removes the decimal point)
            else:
                return "{:.2f}".format(num) #If not an integer, return as string formatted with two decimal places
        except ValueError:
            return '' #if conversion fails, return empty string
            
    def detect_header_row(self, file_path, header_indicator="Game"):
    #Handles automatically detecting header rows by scanning all rows for Wager Audit file/Op Wager Config Sheet
        if file_path.endswith('.xlsx'): #Read Excel file
            wager_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            wager_data = wager_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

        elif file_path.endswith('.csv'): #Handles csv files differently
            rows = [] #Empty list to store rows
            with open(file_path, 'r', encoding='ISO-8859-1') as f: #DEBUG to print first 5 lines from CSV/Wager Audit file:
                reader = csv.reader(f)
                print("\nDEBUG FOR WAGER AUDIT: Preview of raw CSV rows:")

                for i, row in enumerate(reader): #Iterate over each row
                    standardized_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]
                    
                    if i < 5: #DEBUG: print standardized row for first 5 rows
                        print(f"Line {i}: {standardized_row}")
                    rows.append(standardized_row) #Append normalized row to the list of rows

            #Convert rows to DataFrame after reading rows, replace empty strings, None values with NaN for easier handling
            wager_data = pd.DataFrame(rows).replace(['', None], np.nan)
        else:
            raise ValueError("Unsupported file format. Only csv and Excel files are supported.") #Raise error for incorrect file formats
                        
        for idx, row in wager_data.iterrows(): #Iterate through each row, convert all values to string, strip spaces
            row_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            print(f"Checking row {idx}: {row_values}") #DEBUG to print specific header rows it's detecting

            #Check if 'Game' is a part of any column names in this row
            if any(header_indicator in value for value in row_values):
                print(f"Header row detected at index {idx}")

                new_header = wager_data.iloc[idx] #Grab header row use it as new column names
                wager_data = wager_data[(idx + 1):].copy() #Drop all rows above header, keep data rows below header
                wager_data.columns = new_header #Assign new header row to the DataFrame columns
                wager_data.columns = wager_data.columns.astype(str).str.replace('\n', ' ', regex=False).str.strip() #Clean column names
                wager_data = wager_data.loc[:, ~wager_data.columns.duplicated()] #Remove duplicate column names

                return idx
            
        print("No matching header row found.")
        return None

    def detect_version_row(self, file_path, header_version_indicator="Jurisdiction"):
        #Handles automatically detecting header rows by scanning all rows for Agile PLM Report/Op Game List
        if file_path.endswith('.xlsx'): #Read Excel file
            version_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            version_data = version_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

        elif file_path.endswith('.csv'): #Handles csv files differently
            rows = [] #Empty list to store rows
            with open(file_path, 'r', encoding='ISO-8859-1') as f: #DEBUG to print first 5 lines from CSV/VERSION REPORT:
                reader = csv.reader(f)
                print("\nDEBUG FOR GAME/MATH VERSIONS: Preview of raw CSV rows:")

                for i, row in enumerate(reader): #Iterate over each row
                    standardizedversion_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]
                    
                    if i < 5: #DEBUG: print standardized row for first 5 rows
                        print(f"Line {i}: {standardizedversion_row}")

                    rows.append(standardizedversion_row) #Append normalized row to the list of rows

            #Convert rows to DataFrame after reading rows, replace empty strings, None values with NaN for easier handling
            version_data = pd.DataFrame(rows).replace(['', None], np.nan)
        else:
            raise ValueError("Unsupported file format. Only csv and Excel files are supported.") #Raise error for incorrect file formats
                        
        for idx, row in version_data.iterrows(): #Iterate through each row, convert all values to string, strip spaces
            versionrow_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            lowered_values = [val.lower() for val in versionrow_values]

            #Check if 'Jurisdiction' is a part of any column names in this row
            if any(header_version_indicator.lower() in val for val in lowered_values):
                print(f"Header row detected at index {idx}")
                return idx
            
        print("No matching header row found.")
        return None

    def compare_files(self, file_path):
            #Checks if all required files are missing
            if not all([self.wagerauditStaging_path, self.wagerauditProduction_path, self.operatorsheet_path, self.opgamelistStaging_report_path, self.opgamelistProduction_report_path, self.agilereport_path]):
                messagebox.showerror("Error!", "Please upload all required files to proceed.") #Show error if any files are missing
                return False #Stop further execution if files are incomplete
            
            all_valid = True #Set the validation flag to True if all files are present and proceed with processing
            
            #Step 1: process Wager Staging/Production Audit csv Files and Operator Wager Config Excel Sheet
            try:
                #Checks required columns are present in both files
                wageraudit_columns = ["Everi Game ID", "RTP MAX", "Denom", "Line Selection", "Bet Multiplier Selection", "Default Denom", "Default Line", 
                                      "Default Bet Multiplier", "Default Bet", "Min Bet", "Max Bet"]
            
                operatorsheet_columns = ["Game", "RTP%", "Denom Selection", "Line/Ways Selection", "Bet Multiplier Selection", "Default Denom Selection", "Default Line/Ways", 
                                        "Default Bet Multiplier", "Total Default Bet", "Min Bet", "Max Bet"]

                #Defining column mapping for wager audit manually so that names match data
                column_mapping_wager = {
                    "Everi Game ID": "Game",
                    "RTP MAX": "RTP%",
                    "Denom": "Denom Selection",
                    "Line Selection": "Line/Ways Selection",
                    "Bet Multiplier Selection": "Bet Multiplier Selection",
                    "Default Denom": "Default Denom Selection",
                    "Default Line": "Default Line/Ways",
                    "Default Bet Multiplier": "Default Bet Multiplier",
                    "Default Bet": "Total Default Bet",
                    "Min Bet": "Min Bet",
                    "Max Bet": "Max Bet"
                }

                #Detect the header rows for files automatically finding column names
                wagerauditStaging_header_row = self.detect_header_row(self.wagerauditStaging_path, header_indicator="Everi Game ID")
                wagerauditProduction_header_row = self.detect_header_row(self.wagerauditProduction_path, header_indicator="Everi Game ID")
                operatorsheet_header_row = self.detect_header_row(self.operatorsheet_path, header_indicator="Game")

                #Throws an error if no valid header rows are found in the files
                if wagerauditStaging_header_row is None or wagerauditProduction_header_row is None or operatorsheet_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows the Wager Audit csv file and the Operator Wager Config Excel Sheet.")
                    return False
            
                #Read full files, skipping the detected header rows
                if self.wagerauditStaging_path.endswith('.csv'):
                    wagerauditStaging_file = pd.read_csv(self.wagerauditStaging_path, skiprows=wagerauditStaging_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Wager Staging Audit csv file. Only .csv is supported.") #Raise error if incorrect file type is selected
                
                #Read full files, skipping the detected header rows
                if self.wagerauditProduction_path.endswith('.csv'):
                    wagerauditProduction_file = pd.read_csv(self.wagerauditProduction_path, skiprows=wagerauditProduction_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Wager Production Audit csv file. Only .csv is supported.") #Raise error if incorrect file type is selected

                if self.operatorsheet_path.endswith('.xlsx'):
                    operatorsheet_file = pd.read_excel(self.operatorsheet_path, header=operatorsheet_header_row, engine='openpyxl') #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Operator Wager Config Excel Sheet. Only .xlsx is supported.") #Raise error if incorrect file type is selected
                               
                #Normalize column names, strip spaces
                wagerauditStaging_file.columns = wagerauditStaging_file.columns.astype(str).str.strip()
                wagerauditProduction_file.columns = wagerauditProduction_file.columns.astype(str).str.strip()
                operatorsheet_file.columns = operatorsheet_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                wagerauditStaging_file = wagerauditStaging_file[wageraudit_columns]
                wagerauditProduction_file = wagerauditProduction_file[wageraudit_columns]
                operatorsheet_file = operatorsheet_file[operatorsheet_columns]

                #Identify if expected columns are missing
                missing_wagerauditStaging_columns = [col for col in wageraudit_columns if col not in wagerauditStaging_file.columns]
                missing_wagerauditProduction_columns = [col for col in wageraudit_columns if col not in wagerauditProduction_file.columns]
                missing_operatorsheet_columns = [col for col in operatorsheet_columns if col not in operatorsheet_file.columns]

                #Checks for missing columns
                if missing_wagerauditStaging_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Wager Staging Audit csv file: {', '.join(missing_wagerauditStaging_columns)}")
                    return False
                if missing_wagerauditProduction_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Wager Production Audit csv file: {', '.join(missing_wagerauditProduction_columns)}")
                    return False
                if missing_operatorsheet_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Operator Wager Config Excel Sheet: {', '.join(missing_operatorsheet_columns)}")
                    return False
                              
                #Renames columns to match column mapping
                try:
                    wagerauditStaging_file = wagerauditStaging_file.rename(columns=column_mapping_wager)
                    wagerauditProduction_file = wagerauditProduction_file.rename(columns=column_mapping_wager)
                    operatorsheet_file = operatorsheet_file.rename(columns=column_mapping_wager)
                except Exception as e:
                    messagebox.showerror("Error in column_mapping_wager", str(e))
                    return False
                
                #Handles all missing columns by adding them with NaN values to both DataFrames
                for col in column_mapping_wager.values():
                    if col not in wagerauditStaging_file.columns:
                        wagerauditStaging_file[col] = pd.NA
                    if col not in wagerauditProduction_file.columns:
                        wagerauditProduction_file[col] = pd.NA
                    if col not in operatorsheet_file.columns:
                        operatorsheet_file[col] = pd.NA

                #Applies normalization to columns
                wagerauditStaging_file['Game'] = wagerauditStaging_file['Game'].apply(self.normalize_name)
                wagerauditProduction_file['Game'] = wagerauditProduction_file['Game'].apply(self.normalize_name)
                operatorsheet_file['Game'] = operatorsheet_file['Game'].apply(self.normalize_name)
               
                #Fill NaN values with 'N/A' for consistency during comparison/export
                wagerauditStaging_file = wagerauditStaging_file.fillna('N/A')
                wagerauditProduction_file = wagerauditProduction_file.fillna('N/A')
                operatorsheet_file = operatorsheet_file.fillna('N/A')

                #Sorts Game columns alphabetically in all DataFrames
                wagerauditStaging_file = wagerauditStaging_file.sort_values(by='Game', ascending=True)
                wagerauditProduction_file = wagerauditProduction_file.sort_values(by='Game', ascending=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True)

                #Removes duplicates in DataFrames to ensure it only appears once
                wagerauditStaging_file = wagerauditStaging_file.drop_duplicates(subset='Game')
                wagerauditProduction_file = wagerauditProduction_file.drop_duplicates(subset='Game')
                operatorsheet_file = operatorsheet_file.drop_duplicates(subset='Game')
               
                #Ensures DataFrames have only matching Game values
                common_games_wager = (
                    set(wagerauditStaging_file['Game']) &
                    set(wagerauditProduction_file['Game']) &
                    set(operatorsheet_file['Game'])
                )

                #Get sets of Game Titles from each file
                games_wagerauditStaging_file = set(wagerauditStaging_file['Game'])
                games_wagerauditProduction_file = set(wagerauditProduction_file['Game'])
                games_operatorsheet_file = set(operatorsheet_file['Game'])

                #Union of all Game Titles across all three files
                all_games = games_wagerauditStaging_file | games_wagerauditProduction_file | games_operatorsheet_file

                allmissing_games = [] #empty list to collect missing Game Titles

                #Loop through all Game Titles to see which are missing
                for game in all_games:
                    missing_in = []
                    if game not in games_wagerauditStaging_file:
                        missing_in.append('Missing in Wager Staging Audit csv file')
                    if game not in games_wagerauditProduction_file:
                        missing_in.append('Missing in Wager Production Audit csv file')
                    if game not in games_operatorsheet_file:
                        missing_in.append('Missing in Operator Wager Config Excel file')

                    #Append one row per Game Title with combined missing info
                    if missing_in:
                        combined_status = ', '.join(missing_in)
                        allmissing_games.append({
                            'Game': game,
                            'Status': combined_status
                        })

                #Convert missing Game Titles list of dicts into a DataFrame and sort it for sheet 8
                missing_games_wager = pd.DataFrame(allmissing_games).sort_values(by='Game').reset_index(drop=True)

                #Filer rows based on common Game Titles in all three files
                wagerauditStaging_file = wagerauditStaging_file[wagerauditStaging_file['Game'].isin(common_games_wager)]
                wagerauditProduction_file = wagerauditProduction_file[wagerauditProduction_file['Game'].isin(common_games_wager)]
                operatorsheet_file = operatorsheet_file[operatorsheet_file['Game'].isin(common_games_wager)]

                #Sort both DataFrames by 'Game' column and reset index to maintain alignment
                wagerauditStaging_file = wagerauditStaging_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                wagerauditProduction_file = wagerauditProduction_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for sheet 3 to hold side-by-side columns for comparison
                audit_results_wagers = pd.DataFrame()

                #Single loop to handle renamed columns to normalize values and add columns side by side
                for wager_column in wagerauditStaging_file.columns:
                    wagerauditStaging_file[wager_column] = wagerauditStaging_file[wager_column].apply(self.normalize_value) #Normalize Wager Staging Audit file columns

                    #Checks if column exists in Wager Production Audit file
                    if wager_column in wagerauditProduction_file.columns:
                        wagerauditProduction_file[wager_column] = wagerauditProduction_file[wager_column].apply(self.normalize_value) #Normalize Wager Production Audit file

                    #Checks if column exists in operatorsheet_file
                    if wager_column in operatorsheet_file.columns:
                        operatorsheet_file[wager_column] = operatorsheet_file[wager_column].apply(self.normalize_value) #Normalize operator wager config sheet columns

                    if (
                        wager_column in wagerauditStaging_file.columns and
                        wager_column in wagerauditProduction_file.columns and
                        wager_column in operatorsheet_file.columns
                    ):
                        #Side by side columns from all sheets to the DataFrame
                        audit_results_wagers[f"{wager_column}\n(Wager Staging Audit File): "] = wagerauditStaging_file[wager_column]
                        audit_results_wagers[f"{wager_column}\n(Wager Production Audit File): "] = wagerauditProduction_file[wager_column]
                        audit_results_wagers[f"{wager_column}\n({Path(self.operatorsheet_path).stem[:31]}): "] = operatorsheet_file[wager_column]
                    else:
                        if wager_column not in wagerauditStaging_file.columns:
                            print(f"'{wager_column}' not found in the Wager Staging Audit csv File.")
                        if wager_column not in wagerauditProduction_file.columns:
                            print(f"'{wager_column}' not found in the Wager Production Audit csv File.")
                        if wager_column not in operatorsheet_file.columns:
                            print(f"'{wager_column}' not found in the Operator Wager Config Excel Sheet.")

                        #Collect missing games from all files for sheet 7
                        missing_games_wager = pd.concat([missing_games_wager, pd.DataFrame({'Missing Games': [wager_column]})], ignore_index=True)

            except Exception as e:
                all_valid = False
                print(f"Error caught in except block: {e}")
                messagebox.showerror("Error", f"An error has occured for the Wager Staging Audit csv file, Wager Production Audit csv file and Operator Wager Config Excel Sheet: {str(e)}")
                return False
            
            #Step 2: Process Operator Staging/Production GameList csv Report and Agile PLM Excel Report
            try:
                #Checks required columns are present in all files
                opgamelist_columns = ["jurisdictionId", "gameId", "mathVersion", "Version"]
                agilereport_columns = ["Jurisdiction", "GameName", "Math Version", "Latest Software Version"]

                #Defining column mapping for version audit manually so that names match data
                column_mapping_versions = {
                    "jurisdictionId": "Jurisdiction",
                    "gameId": "GameName",
                    "mathVersion": "Math Version",
                    "Version": "Latest Software Version"
                }

                #Detect the header rows for files automatically finding column names
                opgamelistStaging_header_row = self.detect_version_row(self.opgamelistStaging_report_path, header_version_indicator="jurisdictionId")
                opgamelistProduction_header_row = self.detect_version_row(self.opgamelistProduction_report_path, header_version_indicator="jurisdictionId")
                agilereport_header_row = self.detect_version_row(self.agilereport_path, header_version_indicator="Jurisdiction")

                #Throws an error if no valid header rows are found in the files
                if opgamelistStaging_header_row is None or opgamelistProduction_header_row is None or agilereport_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows for the Operator Staging GameList csv Report, Operator Production GameList csv Report and Agile PLM Excel Report.")
                    return False

                #Read full files, skipping the detected header rows
                if self.opgamelistStaging_report_path.endswith('.csv'):
                    opgamelistStaging_file = pd.read_csv(self.opgamelistStaging_report_path, skiprows=opgamelistStaging_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Operator GameList csv Report. Only .csv is supported.") #Raise error if incorrect file type is selected
                
                if self.opgamelistProduction_report_path.endswith('.csv'):
                    opgamelistProduction_file = pd.read_csv(self.opgamelistProduction_report_path, skiprows=opgamelistProduction_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Operator GameList csv Report. Only .csv is supported.") #Raise error if incorrect file type is selected

                if self.agilereport_path.endswith('.xlsx'):
                    agilereport_file = pd.read_excel(self.agilereport_path, header=agilereport_header_row, engine='openpyxl', dtype=str) #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Agile PLM Excel Report. Only .xlsx is supported.") #Raise error if incorrect file type is selected

                #Normalize column names, strip spaces
                opgamelistStaging_file.columns = opgamelistStaging_file.columns.astype(str).str.strip()
                opgamelistProduction_file.columns = opgamelistProduction_file.columns.astype(str).str.strip()
                agilereport_file.columns = agilereport_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                opgamelistStaging_file = opgamelistStaging_file[opgamelist_columns]
                opgamelistProduction_file = opgamelistProduction_file[opgamelist_columns]
                agilereport_file = agilereport_file[agilereport_columns]

                #Renames columns to match column mapping; renames 'GameName' column to 'Game' for consistency
                try:
                    opgamelistStaging_file = opgamelistStaging_file.rename(columns=column_mapping_versions)
                    opgamelistProduction_file = opgamelistProduction_file.rename(columns=column_mapping_versions)
                    agilereport_file = agilereport_file.rename(columns=column_mapping_versions)

                    if 'GameName' in opgamelistStaging_file.columns:
                        opgamelistStaging_file = opgamelistStaging_file.rename(columns={'GameName': 'Game'})
                    if 'GameName' in opgamelistProduction_file.columns:
                        opgamelistProduction_file = opgamelistProduction_file.rename(columns={'GameName': 'Game'})
                    if 'GameName' in agilereport_file.columns:
                        agilereport_file = agilereport_file.rename(columns={'GameName': 'Game'})

                except Exception as e:
                    messagebox.showerror("Error in column_mapping_versions", str(e))
                    return False
                
                #List of column names to store
                final_column_names = list(opgamelistStaging_file.columns)

                #Identify if expected columns are missing
                missing_opgamelistStaging_columns = [col for col in final_column_names if col not in opgamelistStaging_file.columns]
                missing_opgamelistProduction_columns = [col for col in final_column_names if col not in opgamelistProduction_file.columns]
                missing_agilereport_columns = [col for col in final_column_names if col not in agilereport_file.columns]

                #Checks for missing columns
                if missing_opgamelistStaging_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Operator Staging GameList csv Report: {', '.join(missing_opgamelistStaging_columns)}")
                    return False
                
                if missing_opgamelistProduction_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Operator Production GameList csv Report: {', '.join(missing_opgamelistProduction_columns)}")
                    return False
                
                if missing_agilereport_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Agile PLM Excel Report: {', '.join(missing_agilereport_columns)}")
                    return False

                #Handles all missing columns by adding them with NaN values to both DataFrames
                for col in final_column_names:
                    if col not in opgamelistStaging_file.columns:
                        opgamelistStaging_file[col] = np.nan
                    if col not in opgamelistProduction_file.columns:
                        opgamelistProduction_file[col] = np.nan
                    if col not in agilereport_file.columns:
                        agilereport_file[col] = np.nan

                #Applies normalization to columns
                opgamelistStaging_file['Game'] = opgamelistStaging_file['Game'].apply(self.normalize_name)
                opgamelistProduction_file['Game'] = opgamelistProduction_file['Game'].apply(self.normalize_name)
                agilereport_file['Game'] = agilereport_file['Game'].apply(self.normalize_name)

                #Fill NaN values with 'N/A' for consistency during comparison/export
                opgamelistStaging_file = opgamelistStaging_file.fillna('N/A')
                opgamelistProduction_file = opgamelistProduction_file.fillna('N/A')
                agilereport_file = agilereport_file.fillna('N/A')

                #Removes duplicates in DataFrames to ensure it only appears once
                opgamelistStaging_file = opgamelistStaging_file.drop_duplicates(subset='Game')
                opgamelistProduction_file = opgamelistProduction_file.drop_duplicates(subset='Game')
                agilereport_file = agilereport_file.drop_duplicates(subset='Game', keep='last') #keeps last listed version as it is the latest approved per the Agile PLM Report specifically

                #Sorts 'Game' column alphabetically in DataFrames
                opgamelistStaging_file = opgamelistStaging_file.sort_values(by='Game', ascending=True)
                opgamelistProduction_file = opgamelistProduction_file.sort_values(by='Game', ascending=True)
                agilereport_file = agilereport_file.sort_values(by='Game', ascending=True)

                #Ensures DataFrames have only matching Game values
                common_games_version = (
                    set(opgamelistStaging_file['Game']) &
                    set(opgamelistProduction_file['Game']) &
                    set(agilereport_file['Game'])
                )

                #Find missing games from each sheet
                games_opgamelistStaging_file = set(opgamelistStaging_file['Game'])
                games_opgamelistProduction_file = set(opgamelistProduction_file['Game'])
                games_agilereport_file = set(agilereport_file['Game'])

                #Union for all game titles across all three files
                all_games_versions = games_opgamelistStaging_file | games_opgamelistProduction_file | games_agilereport_file

                allmissing_gameVersions = [] #empty list to collect missing Game Titles

                #Loop through all Game Titles to see which are missing
                for gameVersions in all_games_versions:
                    missingVersions_in = []
                    if gameVersions not in games_opgamelistStaging_file:
                        missingVersions_in.append('Missing in Operator Staging GameList csv Report')
                    if gameVersions not in games_opgamelistProduction_file:
                        missingVersions_in.append('Missing in Operator Production GameList csv Report')
                    if gameVersions not in games_agilereport_file:
                        missingVersions_in.append('Missing in Agile PLM Excel Report')

                    #Append one row per Game Title with combined missing info
                    if missingVersions_in:
                        combinedVersions_status = ', '.join(missingVersions_in)
                        allmissing_gameVersions.append({
                            'Game': gameVersions,
                            'Status': combinedVersions_status
                        })

                #Convert missing Game Titles list of dicts into a DataFrame and sort it for sheet 8
                missing_games_versions = pd.DataFrame(allmissing_gameVersions).sort_values(by='Game').reset_index(drop=True)

                #Filer rows based on common games
                opgamelistStaging_file = opgamelistStaging_file[opgamelistStaging_file['Game'].isin(common_games_version)]
                opgamelistProduction_file = opgamelistProduction_file[opgamelistProduction_file['Game'].isin(common_games_version)]
                agilereport_file = agilereport_file[agilereport_file['Game'].isin(common_games_version)]

                #Sort DataFrames by 'Game' column and reset index to maintain alignment
                opgamelistStaging_file = opgamelistStaging_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                opgamelistProduction_file = opgamelistProduction_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                agilereport_file = agilereport_file.sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for sheet 6 to hold side-by-side columns for comparison
                audit_results_versions = pd.DataFrame()

                #Single loop to handle renamed columns and add columns side by side
                for version_column in opgamelistStaging_file.columns:
                    if (
                        version_column in opgamelistStaging_file.columns and
                        version_column in opgamelistProduction_file.columns and
                        version_column in agilereport_file.columns
                    ):
                        #Side by side columns from all sheets to the DataFrame
                        audit_results_versions[f"{version_column}\n(Operator Staging GameList Report): "] = opgamelistStaging_file[version_column]
                        audit_results_versions[f"{version_column}\n(Operator Production GameList Report): "] = opgamelistProduction_file[version_column]
                        audit_results_versions[f"{version_column}\n(Agile PLM Report): "] = agilereport_file[version_column]
                    else:
                        if version_column not in opgamelistStaging_file.columns:
                            print(f"'{version_column}' not found in the Operator Staging GameList Report.")
                        if version_column not in opgamelistProduction_file.columns:
                            print(f"'{version_column}' not found in the Operator Production GameList Report.")
                        if version_column not in agilereport_file.columns:
                            print(f"'{version_column}' not found in the Agile PLM Report.")

                        #Collect missing games from agile plm report and op game list for sheet 8
                        missing_games_versions = pd.concat([missing_games_versions, pd.DataFrame({'Missing Games': [version_column]})], ignore_index=True)

            except Exception as e:
                all_valid = False
                messagebox.showerror("Error!", f"An error has occured for the Operator Staging GameList csv Report, Operator Production GameList csv Report, and Agile PLM Excel Report: {str(e)}")
                return False
            
            #Combine missing games from wager audit and versions audit
            combined_missing_games = pd.concat([missing_games_wager, missing_games_versions], ignore_index=True)
            
            #If all files are processed successfully and True, proceed with Excel writing
            if all_valid:
                try:
                    #Write to excel with formatting
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        #Write to Excel with sheet names (based on selected file paths) truncated to 31 characters
                        wagerauditStaging_file.to_excel(writer, sheet_name=Path(self.wagerauditStaging_path).stem[:31], index=False) #Wager Staging Audit csv file data on sheet 1
                        wagerauditProduction_file.to_excel(writer, sheet_name=Path(self.wagerauditProduction_path).stem[:31], index=False) #Wager Production Audit csv file data on sheet 2
                        operatorsheet_file.to_excel(writer, sheet_name=Path(self.operatorsheet_path).stem[:31], index=False) #Op Wager Config Sheet data on sheet 3
                        audit_results_wagers.to_excel(writer, sheet_name='Wager Audit Comparison Results', index=False) #Wager Audit Comparison Results with side by side comparison on sheet 4
                        opgamelistStaging_file.to_excel(writer, sheet_name=Path(self.opgamelistStaging_report_path).stem[:31], index=False) #Op Staging GameList csv Report data on sheet 5
                        opgamelistProduction_file.to_excel(writer, sheet_name=Path(self.opgamelistProduction_report_path).stem[:31], index=False) #Op Production GameList csv Report data on sheet 6
                        agilereport_file.to_excel(writer, sheet_name=Path(self.agilereport_path).stem[:31], index=False) #Agile PLM Report data on sheet 7
                        audit_results_versions.to_excel(writer, sheet_name='GameVersion Audit Results', index=False) #GameVersion Audit Results with side by side comparison on sheet 8
                        combined_missing_games.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games from all files on sheet 9

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF6F6F'}) #Red format highlights cells red when there's a mismatch on the Wager Audit Comparison Results

                        #Loop & apply formats to all sheets
                        for df, sheet_name in [
                            (wagerauditStaging_file, Path(self.wagerauditStaging_path).stem[:31]),
                            (wagerauditProduction_file, Path(self.wagerauditProduction_path).stem[:31]),
                            (operatorsheet_file, Path(self.operatorsheet_path).stem[:31]),
                            (audit_results_wagers, 'Wager Audit Comparison Results'),
                            (opgamelistStaging_file, Path(self.opgamelistStaging_report_path).stem[:31]),
                            (opgamelistProduction_file, Path(self.opgamelistProduction_report_path).stem[:31]),
                            (agilereport_file, Path(self.agilereport_path).stem[:31]),
                            (audit_results_versions, 'GameVersion Audit Results'),
                            (combined_missing_games, 'Missing Games')
                        ]:
                            worksheet = writer.sheets[sheet_name]

                            #Header row formatting and auto-adjust column widths
                            for col_num, column_name in enumerate(df.columns):
                                worksheet.write(0, col_num, column_name, header_format)

                                #Auto-adjust column width to fit contents by calculating optimal column widths based on header/data length
                                if df[column_name].notna().any():
                                    max_val_len = df[column_name].astype(str).map(len).max()
                                else:
                                    max_val_len = 0

                                max_len = max(max_val_len, len(column_name))
                                worksheet.set_column(col_num, col_num, max_len + 2) #add padding

                            worksheet.autofilter(0, 0, 0, len(df.columns) - 1) #Add filter to header row
                            worksheet.freeze_panes(1, 0) #Freeze top row to keep headers visible when scrolling

                            #Write all data cells w/border formatting
                            for row in range(1, len(df) + 1):
                                for col in range(len(df.columns)):
                                    val = df.iat[row - 1, col]
                                    if pd.isna(val) or val in [float('inf'), float('-inf')]:
                                        worksheet.write(row, col, "", cell_format)
                                    else:
                                        worksheet.write(row, col, val, cell_format)
                                                              
                            #Highlights mismatched values in Wager Audit Comparison Results and GameVersion Audit Results
                            if df is audit_results_wagers or df is audit_results_versions:
                                normalize = df is audit_results_wagers #only normalize values for wager audit

                                #Iterates through rows/columns to apply formatting for mismatches
                                for row in range(1, len(df) + 1):
                                    for col_idx in range(0, len(df.columns), 3): #Iterates over columns
                                        try:
                                            val1 = df.iloc[row - 1, col_idx]
                                            val2 = df.iloc[row - 1, col_idx + 1]
                                            val3 = df.iloc[row - 1, col_idx + 2]

                                            #Excluding column name 'Jurisdiction' from being highlighted red due to inconsistencies on Agile PLM report
                                            column_name = df.columns[col_idx]
                                            if 'Jurisdiction' in column_name:
                                                continue

                                            if normalize:
                                                #Normalize values if necessary
                                                if isinstance(val1, (int, float, str)):
                                                    val1 = self.normalize_value(val1)
                                                if isinstance(val2, (int, float, str)):
                                                    val2 = self.normalize_value(val2)
                                                if isinstance(val3, (int, float, str)):
                                                    val3 = self.normalize_value(val3)

                                            #Convert NaN to empty strings for comparison
                                            safe_val1 = "" if pd.isna(val1) else val1
                                            safe_val2 = "" if pd.isna(val2) else val2
                                            safe_val3 = "" if pd.isna(val3) else val3

                                            #Checks if all values are equal and if so write normally
                                            if safe_val1 == safe_val2 == safe_val3:
                                                worksheet.write(row, col_idx, safe_val1)
                                                worksheet.write(row, col_idx + 1, safe_val2)
                                                worksheet.write(row, col_idx + 2, safe_val3)
                                            else:
                                                #If not all values are equal, highlight column with mismatch only
                                                if safe_val1 == safe_val2: #If val1 and val2 match, val3 is different
                                                    worksheet.write(row, col_idx, safe_val1)
                                                    worksheet.write(row, col_idx + 1, safe_val2)
                                                    worksheet.write(row, col_idx + 2, safe_val3, red_format)

                                                elif safe_val1 == safe_val3: #If val1 and val3 match, val2 is different
                                                    worksheet.write(row, col_idx, safe_val1)
                                                    worksheet.write(row, col_idx + 1, safe_val2, red_format)
                                                    worksheet.write(row, col_idx + 2, safe_val3)

                                                elif safe_val1 == safe_val3: #If val2 and val3 match, val1 is different
                                                    worksheet.write(row, col_idx, safe_val1, red_format)
                                                    worksheet.write(row, col_idx + 1, safe_val2)
                                                    worksheet.write(row, col_idx + 2, safe_val3)

                                                else: #If all three values are different then highlight all
                                                    worksheet.write(row, col_idx, safe_val1, red_format)
                                                    worksheet.write(row, col_idx + 1, safe_val2, red_format)
                                                    worksheet.write(row, col_idx + 2, safe_val3, red_format)                                                                                                    

                                        except Exception as e:
                                            print(f"Error processing row {row}, col_idx {col_idx}: {e}")

                except Exception as e:
                    all_valid = False
                    messagebox.showerror("Error writing to Excel", str(e))
                    return False

            #Success message when results are True and all passes successfully
            if all_valid:
                messagebox.showinfo("Success!", "All files processed successfully and Final Audit Results are complete.")
                return True
            else:
                return False
            

if __name__ =="__main__":
    root = tk.Tk()
    app = CompareFiles(root)
    root.mainloop()
