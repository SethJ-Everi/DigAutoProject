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
from difflib import SequenceMatcher #Import SequenceMatcher for computing similarity between two strings

class FullAuditProgram:
    def __init__(self, master=None):
        self.window = tk.Toplevel(master)
        self.window.title("Wager & Game Version Audit Comparison Tool") #window title
        self.window.configure(bg="#2b2b2b") #set window background color to white

        self.window.protocol("WM_DELETE_WINDOW", self.close_window) #X button will confirm if user wants to close

        self.wagerAudit_Staging_path = "" #path for Wager Staging Audit File
        self.wagerAudit_Production_path = "" #path for Wager Production Audit File
        self.operator_wagerSheet_path = "" #path for Op Wager Config Sheet File
        self.opGameList_stagingReport_path = "" #path for Op Staging GameList Report
        self.opGameList_productionReport_path = "" #path for Op Production GameList Report
        self.agileReport_path = "" #path for the Agile PLM Report

        self.create_widgets() #function for UI components
        self.adjust_window() #function for screen function

        #Default and min size settings
        self.window.geometry("800x600")
        self.window.minsize(800, 600)

    def close_window(self): #Function for cancel confirmation
        confirm = messagebox.askyesno(
            "Exit Wager & Game Version Audit",
            "Are you sure you want to close the Wager & Game Version Audit?"
        )
        if confirm:
            self.window.destroy() #To close this window only
        else:
            messagebox.showinfo(
                "Canceled!",
                "Close cancelled."
            )

    def adjust_window(self):
        #Get the screen's full width/height
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 800
        window_height = 600

        #Calculate position to center the window
        position_top = (screen_height - window_height) // 2
        position_left = (screen_width - window_width) // 2

        #Update the window's geometry to apply size and position
        self.window.geometry(f'{screen_width}x{window_height}+{position_left}+{position_top}')

    def create_widgets(self):
        #Main content frame for all buttons/labels
        content_frame = tk.Frame(self.window, bg="#2b2b2b", height=300)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)

        #Welcome display text and label
        welcome_text = "\nWager & Game Version \nAudit Comparison Tool\n"
        self.welcome_label = tk.Label(content_frame, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='white', bg='#2b2b2b')
        self.welcome_label.pack(pady=10)

        #Container for left/right groups side by side
        group_container = tk.Frame(content_frame, bg="#2b2b2b")
        group_container.pack()

        #Left group (Wager audit files)
        left_group = tk.LabelFrame(group_container, text="Wager Audit Files", font=("TkDefaultFont", 8, "bold"), fg='white', bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
        left_group.pack(side="left", padx=10)

        #Right group (Game/Math version files)
        right_group = tk.LabelFrame(group_container, text="Game & Math Version Files", font=("TkDefaultFont", 8, "bold"), fg='white', bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
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

        #Staging Wager Audit label and upload button
        self.wagerAudit_Staging_label = tk.Label(left_group, text="Select Staging Wager Audit File", **label_style)
        self.wagerAudit_Staging_label.pack(pady=(0, 5))
        self.wagerAudit_Staging_button = tk.Button(left_group, text="Upload Staging Wager Audit File", width=38, command=self.upload_wagerAudit_Staging, **button_style)
        self.wagerAudit_Staging_button.pack(pady=(0, 10))
        self.button_hover_effect(self.wagerAudit_Staging_button) 

        #Production Wager Audit label and upload button
        self.wagerAudit_Production_label = tk.Label(left_group, text="Select Production Wager Audit File", **label_style)
        self.wagerAudit_Production_label.pack(pady=(10, 5))
        self.wagerAudit_Production_button = tk.Button(left_group, text="Upload Production Wager Audit File", width=38, command=self.upload_wagerAudit_Production, **button_style)
        self.wagerAudit_Production_button.pack(pady=(0, 10))
        self.button_hover_effect(self.wagerAudit_Production_button)

        #Operator Wager Config Sheet label and upload button
        self.operator_wagerSheet_label = tk.Label(left_group, text="Select Operator Wager Configuration Sheet", **label_style)
        self.operator_wagerSheet_label.pack(pady=(10, 5))
        self.operator_wagerSheet_button = tk.Button(left_group, text="Upload Operator Wager Configuration Sheet", width=38, command=self.upload_operatorWagerSheet, **button_style)
        self.operator_wagerSheet_button.pack(pady=(0, 10))
        self.button_hover_effect(self.operator_wagerSheet_button)

        #Staging Operator GameList Report label and upload button
        self.opGameList_stagingReport_label = tk.Label(right_group, text="Select Staging Operator GameList Report", **label_style)
        self.opGameList_stagingReport_label.pack(pady=(0, 5))
        self.opGameList_stagingReport_button = tk.Button(right_group, text="Upload Staging Operator GameList Report", width=38, command=self.upload_opGameList_stagingReport, **button_style)
        self.opGameList_stagingReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.opGameList_stagingReport_button)

        #Production Operator GameList Report label and upload button
        self.opGameList_productionReport_label = tk.Label(right_group, text="Select Production Operator GameList Report", **label_style)
        self.opGameList_productionReport_label.pack(pady=(10, 5))
        self.opGameList_productionReport_button = tk.Button(right_group, text="Upload Production Operator GameList Report", width=38, command=self.upload_opGameList_productionReport, **button_style)
        self.opGameList_productionReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.opGameList_productionReport_button)

        #Agile PLM Report label and upload button
        self.agileReport_label = tk.Label(right_group, text="Select Agile PLM Report", **label_style)
        self.agileReport_label.pack(pady=(10, 5))
        self.agileReport_button = tk.Button(right_group, text="Upload Agile PLM Report", width=38, command=self.upload_agileReport, **button_style)
        self.agileReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.agileReport_button)

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
                if button is self.submit_button and button.cget("bg") == "green":
                    button.config(bg="dark green")
                elif button.cget("bg") != "green":
                    button.config(bg=hover_bg)

        def on_leave(e):
            if button.cget("state") == tk.NORMAL:
                if button is self.submit_button and button.cget("bg") == "dark green":
                    button.config(bg="green")
                elif button.cget("bg") == hover_bg:
                    button.config(bg=normal_bg)
                else:
                    button.config(bg="#FF6F6F")

        #Bind hover effect to the button
        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def enable_submit_button(self):
        #Enables the submit button if files are not empty and turns green. Otherwise, displays red if no files are selected and remains disabled
        if all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path, self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
            self.submit_button.config(state=tk.NORMAL, bg='green')
        else:
            self.submit_button.config(state=tk.DISABLED, bg='#FF6F6F')
        self.button_hover_effect(self.submit_button)

    def clear_button(self):
        answer = messagebox.askyesno(
            "Confirm Clear?",
            "Are you sure you want to clear all files selected?"
        )
        if answer:
            #Clear all file paths if yes is selected
            self.wagerAudit_Staging_path = ""
            self.wagerAudit_Production_path = ""
            self.operator_wagerSheet_path = ""
            self.opGameList_stagingReport_path = ""
            self.opGameList_productionReport_path = ""
            self.agileReport_path = ""
            
            #Clear all labels and display red text
            self.wagerAudit_Staging_label.config(text="Select Staging Wager Audit File", fg="#FF6F6F")
            self.wagerAudit_Production_label.config(text="Select Production Wager Audit File", fg="#FF6F6F")
            self.operator_wagerSheet_label.config(text="Select Operator Wager Configuration Sheet", fg="#FF6F6F")
            self.opGameList_stagingReport_label.config(text="Select Staging Operator GameList Report", fg="#FF6F6F")
            self.opGameList_productionReport_label.config(text="Select Production Operator GameList Report", fg="#FF6F6F")
            self.agileReport_label.config(text="Select Agile PLM Report", fg="#FF6F6F")

            #Disable the submit button and turn red
            self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

            #Show message box to user stating cleared files
            messagebox.showinfo("All Files Cleared!",
                                "All uploaded files were cleared. Select new files to upload.")
            
        else: #Show message box to user the clear was canceled
            messagebox.showinfo("Canceled!",
                                "Clear canceled.")
        #Disable the submit button and turn red
        self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

    def upload_wagerAudit_Staging(self):
        self.wagerAudit_Staging_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerAudit_Staging_path: #Checks if a file is selected
            self.wagerAudit_Staging_label.config(text=f"Staging Wager Audit File Uploaded: \n{self.wagerAudit_Staging_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Staging Wager Audit File to proceed.") #Show warning if no staging wager audit file is selected
            self.wagerAudit_Staging_label.config(text="Select Staging Wager Audit File", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.wagerAudit_Staging_path = "" if not self.wagerAudit_Staging_path else self.wagerAudit_Staging_path
        self.enable_submit_button() #Enables submit button after selection

    def upload_wagerAudit_Production(self):
        self.wagerAudit_Production_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerAudit_Production_path: #Checks if a file is selected
            self.wagerAudit_Production_label.config(text=f"Wager Production Audit File Uploaded: \n{self.wagerAudit_Production_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Production Wager Audit File to proceed.") #Show warning if no production wager audit file is selected
            self.wagerAudit_Production_label.config(text="Select Production Wager Audit File", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.wagerAudit_Production_path = "" if not self.wagerAudit_Production_path else self.wagerAudit_Production_path
        self.enable_submit_button() #Enables submit button after selection 

    def upload_operatorWagerSheet(self):
        self.operator_wagerSheet_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file (this is the file type when file is downloaded)

        if self.operator_wagerSheet_path: #Checks if a file is selected
            self.operator_wagerSheet_label.config(text=f"Operator Wager Configuration Sheet Uploaded: \n{self.operator_wagerSheet_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Operator Wager Configuration Sheet to proceed.") #Show warning if no op wager config sheet is selected
            self.operator_wagerSheet_label.config(text="Select Operator Wager Configuration Sheet", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.operator_wagerSheet_path = "" if not self.operator_wagerSheet_path else self.operator_wagerSheet_path
        self.enable_submit_button() #Enables submit button after selection

    def upload_opGameList_stagingReport(self):
        self.opGameList_stagingReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opGameList_stagingReport_path: #Checks if a file is selected
            self.opGameList_stagingReport_label.config(text=f"Staging Operator GameList Report Uploaded: \n{self.opGameList_stagingReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Staging Operator GameList Report to proceed.") #Show warning if no staging op gamelist report is selected 
            self.opGameList_stagingReport_label.config(text="Select Staging Operator GameList Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opGameList_stagingReport_path = "" if not self.opGameList_stagingReport_path else self.opGameList_stagingReport_path
        self.enable_submit_button() #Enables submit button after selection
    
    def upload_opGameList_productionReport(self):
        self.opGameList_productionReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opGameList_productionReport_path: #Checks if a file is selected
            self.opGameList_productionReport_label.config(text=f"Production Operator GameList Report Uploaded: \n{self.opGameList_productionReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Production Operator GameList Report to proceed.") #Show warning if no production op gamelist report is selected 
            self.opGameList_productionReport_label.config(text="Select Production Operator GameList Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opGameList_productionReport_path = "" if not self.opGameList_productionReport_path else self.opGameList_productionReport_path

    def upload_agileReport(self):
        self.agileReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file (this is the file type when file is downloaded from agile power bi)

        if self.agileReport_path: #Checks if a file is selected
            self.agileReport_label.config(text=f"Agile PLM Report Uploaded: \n{self.agileReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Please select Agile PLM Report to proceed.") #Show warning if no agile plm report is selected
            self.agileReport_label.config(text="Select Agile PLM Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.agileReport_path = "" if not self.agileReport_path else self.agileReport_path
        self.enable_submit_button() #Enables submit button after selection

    def submit_files(self):
        #Checks if all files are uploaded
        if not all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path, self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
            messagebox.showwarning("Incomplete files!", "Please upload all required files before submitting.") #Show warning if not all files were uploaded
            return

        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            parent=self.window,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #file types filter
            title="File Save Location" #dialog title
        )

        if not file_path:
            messagebox.showinfo("Missing File Path!", "Select file path to save Wager & Game Version Audit Results and try again.") #Show cancelled message if no save file path was selected
            self.enable_submit_button() #Enables submit button
            return

        #Message box to confirm user selected files for submission and allows user to hit cancel if needed to re-upload files
        if messagebox.askyesno("Confirm Submit", "Are you sure you want to submit files for comparison?"):
            try:
                result = self.compare_files(file_path) #Call the function to compare files and save
                if result:
                    messagebox.showinfo("Audit Results Saved!", f"Wager & Game Version Audit Results successfully saved at: {file_path}.") #Success message and show user save location
                else:
                    messagebox.showerror("Error!", "Failed to save file. Please check the correct file formats (.csv or .xlsx) were submitted and try again.") #Show failure message if results fail
            except Exception as e:
                messagebox.showerror("Error!", f"Error occurred during export: {str(e)}") #Show error if there's an exception while saving files
        else:
            messagebox.showinfo("Canceled!", "File submission canceled. Please upload all required files to submit and try again.") #Display cancel message if user hits cancel

        self.enable_submit_button() #Resets submit button to it's default state after handling success, cancellation, or missing file path

    def normalize_name(self, name):
        #Standardize game name column; convert to lowercase, removes all spaces, removes apostrophes
        if isinstance(name, str):
            name = unicodedata.normalize('NFKD', name) #Normalize any smart quotes or accents (Ex: Jack O'Lantern Jackpots)
            name = re.sub(r'(?<!^)(?=[A-Z][a-z])', ' ', name) #Split only before capital letters followed by lowercase (to avoid splitting acronyms)
            name = name.replace('_', ' ') #Remove underscores and adds a space (specific for postfix games)
            name = re.sub(r"[’';:]", '', name) #Remove straight and curly apostrophes using regex
            name = re.sub(r'\s+', '', name).strip() #Replace multiple spaces with no space, then strip leading/trailing
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
            val = re.sub(r'[$€£,]', '', val).strip() #Remove the currency symbols/commas using regex
            num = float(val) #Convert to float
            if num.is_integer(): #Checks if float is an integer
                return str(int(num)) #If integer, convert to an integer then to a string (removes the decimal point)
            else:
                return "{:.2f}".format(num) #If not an integer, return as string formatted with two decimal places
        except ValueError:
            return '' #if conversion fails, return empty string
            
    def detect_header_row(self, file_path, header_indicator="Game"):
    #Handles automatically detecting header rows by scanning all rows for Wager Files
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
            raise ValueError("Unsupported file format. Only CSV (.csv) and Excel (.xlsx) file types are supported.") #Raise error for incorrect file formats
                        
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
            
        raise ValueError("No matching header row found. Check files to ensure proper files were uploaded and try again.") #Raise error for when headers are not found

    def detect_version_row(self, file_path, header_version_indicator="Jurisdiction", unwanted_keywords=None):
        #Default unwanted words if none are provided
        if unwanted_keywords is None:
            unwanted_keywords = ['Applied filters:', 'is not (Blank)']

        #Handles automatically detecting header rows by scanning all rows for Game/Math Version Files
        if file_path.endswith('.xlsx'): #Read Excel file
            version_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            version_data = version_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

            #DEBUG: Print first 5 rows for inspection
            print("\nDEBUG Excel Files: Preview of first 5 raw rows:")
            print(version_data.head())

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
            raise ValueError("Unsupported file format. Only ('.csv') and ('.xlsx') file types are supported.") #Raise error for incorrect file formats
        
        #Drop row containing unwanted text from Agile PLM Report
        rows_to_drop = [] #Store indexes of rows to drop
        for idx, row in version_data.iterrows():
            row_text = ' '.join(str(cell).lower() for cell in row if isinstance(cell, str)) #Join cell values into a single string, lowercase it
            if any(keyword in row_text for keyword in unwanted_keywords): #Mark row for deletion if found
                print(f"Unwanted verbiage detected on Agile PLM Report at row {idx}, removing that row.")
                rows_to_drop.append(idx)

        #Drop row with unwanted text and reset index
        version_data = version_data.drop(index=rows_to_drop).reset_index(drop=True)

        for idx, row in version_data.iterrows(): #Iterate through each row, convert all values to string, strip spaces
            versionrow_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            lowered_values = [val.lower() for val in versionrow_values]

            #Check if 'Jurisdiction' is a part of any column names in this row
            if any(header_version_indicator.lower() in val for val in lowered_values):
                print(f"Header row detected at index {idx}")
                return idx
            
        raise ValueError("No matching header row found. Check files to ensure proper files were uploaded and try again.") #Raise error for when headers are not found

    def partialMatching_GameNames(self, *Game, min_similarity=0.8):
        #Handles partial Game Name matches and returns true if average similarity across all pairs >= min_similarity (EX: Off The Hook; Good Ol Fishin Hole in the Agile Report vs Good Ol Fishin Hole in Op GameList Reports)
        n = len(Game)
        similarities = []

        for i in range(n):
            for j in range(i + 1, n):
                similarity = SequenceMatcher(None, Game[i], Game[j]).ratio()
                similarities.append(similarity)

        average_similarity = sum(similarities) / len(similarities) if similarities else 0

        lengths = [len(t) for t in Game]
        if max(lengths) / min(lengths) >= 2:
            return False, average_similarity

        return average_similarity >= min_similarity, average_similarity
    
    def matching_GameNames(self, opGameList_StagingReport_gameNames, opGameList_ProductionReport_gameNames, agileReport_gameNames, threshold=85, min_similarity=0.8):
        #Handles Game Name exact + partial matches for all three files
        gameName_matches = []
        matched_opGameList_Staging, matched_opGameList_Production, matched_agileReport = set(), set(), set()

        for t1 in opGameList_StagingReport_gameNames:
            for t2 in opGameList_ProductionReport_gameNames:
                for t3 in agileReport_gameNames:
                    if t1 in matched_opGameList_Staging and t2 in matched_opGameList_Production and t3 in matched_agileReport:
                        continue

                    if t1 == t2 == t3:
                        gameName_matches.append((t1, t2, t3, t3))
                        matched_opGameList_Staging.add(t1)
                        matched_opGameList_Production.add(t2)
                        matched_agileReport.add(t3)
                        print(f"DEBUG: Exact match triggered: '{t1}'") #DEBUG for exact match triggers
                        continue

                    score_t1_t2 = SequenceMatcher(None, t1, t2).ratio() * 100
                    score_t2_t3 = SequenceMatcher(None, t2, t3).ratio() * 100
                    score_t1_t3 = SequenceMatcher(None, t1, t3).ratio() * 100
                    average_score = (score_t1_t2 + score_t2_t3 + score_t1_t3) / 3

                    if average_score >= threshold:
                        gameName_matches.append((t1, t2, t3, t3))
                        matched_opGameList_Staging.add(t1)
                        matched_opGameList_Production.add(t2)
                        matched_agileReport.add(t3)
                        print(f"DEBUG: EXACT match triggered: t1='{t1}', t2='{t2}', t3='{t3}' -> (avg_score={average_score:.2f})") #DEBUG for exact match triggers
                        continue

                    #partial matching logic only if agile plm report differs
                    if t3 != t1:
                        partial_results, average_similarity = self.partialMatching_GameNames(t1, t2, t3, min_similarity=min_similarity)
                        if partial_results:
                            gameName_matches.append((t1, t2, t3, t1))
                            matched_opGameList_Staging.add(t1)
                            matched_opGameList_Production.add(t2)
                            matched_agileReport.add(t3)
                            print(f"DEBUG: Partial match triggered File3: '{t3}' -> '{t1}' -> avg_similarities={average_similarity:.3f}") #DEBUG for partia match triggers

        return gameName_matches #Return the full list of matched game names and their scores

    def compare_files(self, file_path):
            #Checks if all required files are missing
            if not all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path, self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
                messagebox.showerror("Error!", "Please upload all required files to proceed.") #Show error if any files are missing
                return False #Stop further execution if files are incomplete
            
            all_valid = True #Set the validation flag to True if all files are present and proceed with processing
            
            #Step 1: process Wager Staging/Production Audit Files and Operator Wager Config Sheet
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
                wagerauditStaging_header_row = self.detect_header_row(self.wagerAudit_Staging_path, header_indicator="Everi Game ID")
                wagerauditProduction_header_row = self.detect_header_row(self.wagerAudit_Production_path, header_indicator="Everi Game ID")
                operatorsheet_header_row = self.detect_header_row(self.operator_wagerSheet_path, header_indicator="Game")

                #Throws an error if no valid header rows are found in the files
                if wagerauditStaging_header_row is None or wagerauditProduction_header_row is None or operatorsheet_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows in the Staging Wager Audit File, Production Wager Audit File, and Operator Wager Configuration Sheet.")
                    return False
            
                #Read full files, skipping the detected header rows
                if self.wagerAudit_Staging_path.endswith('.csv'):
                    wagerauditStaging_file = pd.read_csv(self.wagerAudit_Staging_path, skiprows=wagerauditStaging_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Staging Wager File. Only '.csv' file type is supported.") #Raise error if incorrect file type is selected
                
                #Read full files, skipping the detected header rows
                if self.wagerAudit_Production_path.endswith('.csv'):
                    wagerauditProduction_file = pd.read_csv(self.wagerAudit_Production_path, skiprows=wagerauditProduction_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Production Wager Audit File. Only '.csv' file type is supported.") #Raise error if incorrect file type is selected

                if self.operator_wagerSheet_path.endswith('.xlsx'):
                    operatorsheet_file = pd.read_excel(self.operator_wagerSheet_path, header=operatorsheet_header_row, engine='openpyxl') #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Operator Wager Configuration Sheet. Only '.xlsx' file type is supported.") #Raise error if incorrect file type is selected
                               
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
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Staging Wager Audit File: {', '.join(missing_wagerauditStaging_columns)}")
                    return False
                if missing_wagerauditProduction_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Production Wager Audit File: {', '.join(missing_wagerauditProduction_columns)}")
                    return False
                if missing_operatorsheet_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Operator Wager Configuration Sheet: {', '.join(missing_operatorsheet_columns)}")
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

                #Get sets of Game Names from each file
                games_wagerauditStaging_file = set(wagerauditStaging_file['Game'])
                games_wagerauditProduction_file = set(wagerauditProduction_file['Game'])
                games_operatorsheet_file = set(operatorsheet_file['Game'])

                #Union of all Game Names across all three files
                all_games = games_wagerauditStaging_file | games_wagerauditProduction_file | games_operatorsheet_file

                allmissing_games = [] #Empty list to collect missing Game Names

                #Loop through all Game Names to see which are missing
                for game in all_games:
                    missing_in = []
                    if game not in games_wagerauditStaging_file:
                        missing_in.append('Missing in Staging Wager Audit File')
                    if game not in games_wagerauditProduction_file:
                        missing_in.append('Missing in Production Wager Audit File')
                    if game not in games_operatorsheet_file:
                        missing_in.append('Missing in Operator Wager Configuration Sheet')

                    #Append one row per Game Name with combined missing info
                    if missing_in:
                        combined_status = ', '.join(missing_in)
                        allmissing_games.append({
                            'Game': game,
                            'Status': combined_status
                        })

                #Convert missing Game Names list of dicts into a DataFrame and sort it for Missing Games Sheet
                missing_games_wager = pd.DataFrame(allmissing_games).sort_values(by='Game').reset_index(drop=True)

                #Filer rows based on common Game Names in all three files
                wagerauditStaging_file = wagerauditStaging_file[wagerauditStaging_file['Game'].isin(common_games_wager)]
                wagerauditProduction_file = wagerauditProduction_file[wagerauditProduction_file['Game'].isin(common_games_wager)]
                operatorsheet_file = operatorsheet_file[operatorsheet_file['Game'].isin(common_games_wager)]

                #Sort both DataFrames by 'Game' column and reset index to maintain alignment
                wagerauditStaging_file = wagerauditStaging_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                wagerauditProduction_file = wagerauditProduction_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for Wager Audit Results to hold side-by-side columns for comparison
                audit_results_wagers = pd.DataFrame()

                #Single loop to handle renamed columns to normalize values and add columns side by side
                for wager_column in wagerauditStaging_file.columns:
                    wagerauditStaging_file[wager_column] = wagerauditStaging_file[wager_column].apply(self.normalize_value) #Normalize Staging Wager Audit File columns

                    #Checks if column exists in Wager Production Audit file
                    if wager_column in wagerauditProduction_file.columns:
                        wagerauditProduction_file[wager_column] = wagerauditProduction_file[wager_column].apply(self.normalize_value) #Normalize Production Wager Audit File

                    #Checks if column exists in operatorsheet_file
                    if wager_column in operatorsheet_file.columns:
                        operatorsheet_file[wager_column] = operatorsheet_file[wager_column].apply(self.normalize_value) #Normalize Operator Wager Config Sheet columns

                    if wager_column == 'Game':
                        continue

                    if (
                        wager_column in wagerauditStaging_file.columns and
                        wager_column in wagerauditProduction_file.columns and
                        wager_column in operatorsheet_file.columns
                    ):
                        #Side by side columns from all sheets to the DataFrame
                        audit_results_wagers[f"{wager_column}\n(Staging Wager Audit File): "] = wagerauditStaging_file[wager_column]
                        audit_results_wagers[f"{wager_column}\n(Production Wager Audit File): "] = wagerauditProduction_file[wager_column]
                        audit_results_wagers[f"{wager_column}\n({Path(self.operator_wagerSheet_path).stem[:31]}): "] = operatorsheet_file[wager_column]
                    else:
                        if wager_column not in wagerauditStaging_file.columns:
                            print(f"'{wager_column}' not found in Staging Wager Audit File.")
                        if wager_column not in wagerauditProduction_file.columns:
                            print(f"'{wager_column}' not found in Production Wager Audit File.")
                        if wager_column not in operatorsheet_file.columns:
                            print(f"'{wager_column}' not found in Operator Wager Configuration Sheet.")

                        #Collect missing games from all files for Missing Games sheet
                        missing_games_wager = pd.concat(
                            [missing_games_wager, pd.DataFrame({'Missing Games': [wager_column]})], ignore_index=True
                            )
                        
                audit_results_wagers['Game'] = wagerauditStaging_file['Game'].values
                cols = list(audit_results_wagers.columns)
                cols.remove('Game')
                cols.insert(0, 'Game')
                audit_results_wagers = audit_results_wagers[cols]

                audit_results_wagers = audit_results_wagers.sort_values(by='Game', ascending=True).reset_index(drop=True)

            except Exception as e:
                all_valid = False
                print(f"Error caught in except block: {e}")
                messagebox.showerror("Error", f"An error has occured for the Staging Wager Audit File, Production Wager Audit File, and Operator Wager Configuration Sheet: {str(e)}")
                return False
            
            #Step 2: Process Staging Operator GameList/Production Operator GameList Reports and Agile PLM Report
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
                opgamelistStaging_header_row = self.detect_version_row(self.opGameList_stagingReport_path, header_version_indicator="jurisdictionId")
                opgamelistProduction_header_row = self.detect_version_row(self.opGameList_productionReport_path, header_version_indicator="jurisdictionId")
                agilereport_header_row = self.detect_version_row(self.agileReport_path, header_version_indicator="Jurisdiction")

                #Throws an error if no valid header rows are found in files
                if opgamelistStaging_header_row is None or opgamelistProduction_header_row is None or agilereport_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows for the Staging Operator GameList Report, Production Operator GameList Report, and Agile PLM Report.")
                    return False

                #Read full files, skipping the detected header rows
                if self.opGameList_stagingReport_path.endswith('.csv'):
                    opgamelistStaging_file = pd.read_csv(self.opGameList_stagingReport_path, skiprows=opgamelistStaging_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Staging Operator GameList Report. Only '.csv' file type is supported.") #Raise error if incorrect file type is selected
                
                if self.opGameList_productionReport_path.endswith('.csv'):
                    opgamelistProduction_file = pd.read_csv(self.opGameList_productionReport_path, skiprows=opgamelistProduction_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Production Operator GameList Report. Only '.csv' file type is supported.") #Raise error if incorrect file type is selected

                if self.agileReport_path.endswith('.xlsx'):
                    agilereport_file = pd.read_excel(self.agileReport_path, header=agilereport_header_row, engine='openpyxl', dtype=str) #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Agile PLM Report. Only '.xlsx' file type is supported.") #Raise error if incorrect file type is selected

                #Normalize column names, strip spaces
                opgamelistStaging_file.columns = opgamelistStaging_file.columns.astype(str).str.strip()
                opgamelistProduction_file.columns = opgamelistProduction_file.columns.astype(str).str.strip()
                agilereport_file.columns = agilereport_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                opgamelistStaging_file = opgamelistStaging_file[opgamelist_columns]
                opgamelistProduction_file = opgamelistProduction_file[opgamelist_columns]
                agilereport_file = agilereport_file[agilereport_columns]

                #Identify if expected columns are missing
                missing_opgamelistStaging_columns = [col for col in opgamelist_columns if col not in opgamelistStaging_file.columns]
                missing_opgamelistProduction_columns = [col for col in opgamelist_columns if col not in opgamelistProduction_file.columns]
                missing_agilereport_columns = [col for col in agilereport_columns if col not in agilereport_file.columns]

                #Checks for missing columns and if missing, program will not continue
                if missing_opgamelistStaging_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Staging Operator GameList Report: {', '.join(missing_opgamelistStaging_columns)}")
                    return False
                
                if missing_opgamelistProduction_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Production Operator GameList Report: {', '.join(missing_opgamelistProduction_columns)}")
                    return False
                
                if missing_agilereport_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Agile PLM Report: {', '.join(missing_agilereport_columns)}")
                    return False

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

                #File labels for labeling on Missing Games sheet
                file_labels = ['Staging Operator GameList Report',
                               'Production Operator GameList Report',
                               'Agile PLM Report']

                #Get Game Name matches from all files
                gameName_matches_versionAudit = self.matching_GameNames(
                    list(opgamelistStaging_file['Game']),
                    list(opgamelistProduction_file['Game']),
                    list(agilereport_file['Game']),
                    threshold=85,
                    min_similarity=0.8
                )

                #Build map for agile plm report game name to op gamelist staging (partial matches)
                agileReport_file_to_opGameList_stagingFile_map = {m[2]: m[0] for m in gameName_matches_versionAudit if m[2] != m[0]}

                #Pre-align agile plm report game names using the mapping above
                agileReport_file_aligned = agilereport_file.copy()
                agileReport_file_aligned['Game'] = agileReport_file_aligned['Game'].apply(
                    lambda t: agileReport_file_to_opGameList_stagingFile_map.get(t, t)
                )

                #Build missing game name sets for detection
                opGameList_stagingFile_set = set(opgamelistStaging_file['Game'])
                opGameList_productionFile_set = set(opgamelistProduction_file['Game'])
                agileReportFile_set = set(agileReport_file_aligned['Game'])

                all_gameNames_union = opGameList_stagingFile_set.union(opGameList_productionFile_set).union(agileReportFile_set)
                allFiles_set = [opGameList_stagingFile_set, opGameList_productionFile_set, agileReportFile_set]

                #Compute missing game names
                allMissing_gameNames_versionAudit = [] #Empty list to collect missing Game Names

                for gameVersions in all_gameNames_union:
                    missingVersions_in =[
                        file_labels[i] for i, gameVersions_sets in enumerate(allFiles_set)
                        if gameVersions not in gameVersions_sets
                    ]
                    #Append one row per Game Name with combined missing info
                    if missingVersions_in:
                        combinedVersions_status = ', '.join(missingVersions_in)
                        allMissing_gameNames_versionAudit.append({
                            'Game': gameVersions,
                            'Status': f'Missing in {combinedVersions_status}'
                        })

                #Convert missing Game Names list of dicts into a DataFrame and sort it for Missing Games sheet
                missing_gameNames_versionAudit = (
                    pd.DataFrame(allMissing_gameNames_versionAudit)
                    .drop_duplicates(subset=['Game', 'Status'])
                    .sort_values(by='Game')
                    .reset_index(drop=True)
                )

                #Collect matched rows for final audit results only if all three files have the same Game Name
                opGameList_stagingFile_matchedGameNames, opGameList_productionFile_matchedGameNames, agileReport_matchedGameNames, gameName_rows = [], [], [], []

                for t1, t2, t3_original, t3_aligned in gameName_matches_versionAudit:
                    row1_staging_df = opgamelistStaging_file.loc[opgamelistStaging_file['Game'] == t1]
                    row2_production_df = opgamelistProduction_file.loc[opgamelistProduction_file['Game'] == t2]
                    row3_agileReport_df = agileReport_file_aligned.loc[agileReport_file_aligned['Game'] == t3_aligned]

                    #Skip if any row is missing
                    if row1_staging_df.empty or row2_production_df.empty or row3_agileReport_df.empty:
                        continue

                    row1_idx_staging = row1_staging_df.iloc[0]
                    row2_idx_production = row2_production_df.iloc[0]
                    row3_idx_agileReport = row3_agileReport_df.iloc[0].copy()

                    if t3_aligned != t3_original:
                        row3_idx_agileReport['Game'] = t1

                    opGameList_stagingFile_matchedGameNames.append(row1_idx_staging.to_dict())
                    opGameList_productionFile_matchedGameNames.append(row2_idx_production.to_dict())
                    agileReport_matchedGameNames.append(row3_idx_agileReport.to_dict())
                    gameName_rows.append({'Game': t1})

                #Build results table
                row1_staging_df = pd.DataFrame(opGameList_stagingFile_matchedGameNames).reset_index(drop=True)
                row2_production_df = pd.DataFrame(opGameList_productionFile_matchedGameNames).reset_index(drop=True)
                row3_agileReport_df = pd.DataFrame(agileReport_matchedGameNames).reset_index(drop=True)
                gameName_rows_df = pd.DataFrame(gameName_rows).drop_duplicates().reset_index(drop=True)
                audit_results_versions = gameName_rows_df.copy()

                all_combined_columns = (
                    set(row1_staging_df.columns)
                    .union(row2_production_df.columns)
                    .union(row3_agileReport_df.columns)
                    - {'Game'}
                )

                #Combine remaining columns with validation
                for col in sorted(all_combined_columns):
                    if col == 'Jurisdiction':
                        continue #Skip adding Jurisdiction to handle separately below
                    if col not in row1_staging_df:
                        raise KeyError(f"{col} not found in 'opgamelistStaging_file' matched rows datasets")
                    if col not in row2_production_df:
                        raise KeyError(f"{col} not found in 'opgamelistProduction_file' matched rows datasets")
                    if col not in row3_agileReport_df:
                        raise KeyError(f"{col} not found in 'agileReport_file_aligned' matched rows datasets")
                    
                    audit_results_versions[f"{col} (Staging Operator GameList Report):"] = row1_staging_df[col].reset_index(drop=True)
                    audit_results_versions[f"{col} (Production Operator GameList Report):"] = row2_production_df[col].reset_index(drop=True)
                    audit_results_versions[f"{col} (Agile PLM Report):"] = row3_agileReport_df[col].reset_index(drop=True)

                #Jurisdiction to only appear once pulled from agile plm report column
                if 'Jurisdiction' in row3_agileReport_df.columns:
                    audit_results_versions['Jurisdiction'] = row3_agileReport_df['Jurisdiction'].reset_index(drop=True)
                    audit_results_versions = audit_results_versions.sort_values(by='Game', ascending=True).reset_index(drop=True)
                    #Rearrange columns putting Jurisdiction before Game
                    cols = list(audit_results_versions.columns)
                    cols.remove('Jurisdiction')
                    gameName_index = cols.index('Game')
                    cols.insert(gameName_index, 'Jurisdiction')
                    audit_results_versions = audit_results_versions[cols]
                else:
                    audit_results_versions = audit_results_versions.sort_values(by='Game', ascending=True).reset_index(drop=True)
                
            except Exception as e:
                all_valid = False
                messagebox.showerror("Error!", f"An error has occured for the Staging Operator GameList Report, Production Operator GameList Report, and Agile PLM Report: {str(e)}")
                return False
            
            #Combine missing games from wager audit and versions audit combined for Missing Games sheet
            combined_missing_games = pd.concat([missing_games_wager, missing_gameNames_versionAudit], ignore_index=True)
            
            #If all files are processed successfully and True, proceed with Excel writing
            if all_valid:
                try:
                    #Write to excel with formatting
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        #Write to Excel with sheet names (based on selected file paths) truncated to 31 characters
                        wagerauditStaging_file.to_excel(writer, sheet_name=Path(self.wagerAudit_Staging_path).stem[:31], index=False) #Staging Wager Audit File raw data on sheet 1
                        wagerauditProduction_file.to_excel(writer, sheet_name=Path(self.wagerAudit_Production_path).stem[:31], index=False) #Production Wager Audit File raw data on sheet 2
                        operatorsheet_file.to_excel(writer, sheet_name=Path(self.operator_wagerSheet_path).stem[:31], index=False) #Op Wager Config Sheet raw data on sheet 3
                        audit_results_wagers.to_excel(writer, sheet_name='Wager Audit Results', index=False) #Wager Audit Results with side by side comparison on sheet 4
                        opgamelistStaging_file.to_excel(writer, sheet_name=Path(self.opGameList_stagingReport_path).stem[:31], index=False) #Staging Op GameList Report raw data on sheet 5
                        opgamelistProduction_file.to_excel(writer, sheet_name=Path(self.opGameList_productionReport_path).stem[:31], index=False) #Production Op GameList Report raw data on sheet 6
                        agilereport_file.to_excel(writer, sheet_name=Path(self.agileReport_path).stem[:31], index=False) #Agile PLM Report raw data on sheet 7
                        audit_results_versions.to_excel(writer, sheet_name='GameVersion Audit Results', index=False) #GameVersion Audit Results with side by side comparison on sheet 8
                        combined_missing_games.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games from all files on sheet 9

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #Borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF6F6F'}) #Red format highlights cells red when there's a mismatch on the Wager Audit Comparison Results

                        #Loop & apply formats to all sheets
                        for df, sheet_name in [
                            (wagerauditStaging_file, Path(self.wagerAudit_Staging_path).stem[:31]),
                            (wagerauditProduction_file, Path(self.wagerAudit_Production_path).stem[:31]),
                            (operatorsheet_file, Path(self.operator_wagerSheet_path).stem[:31]),
                            (audit_results_wagers, 'Wager Audit Results'),
                            (opgamelistStaging_file, Path(self.opGameList_stagingReport_path).stem[:31]),
                            (opgamelistProduction_file, Path(self.opGameList_productionReport_path).stem[:31]),
                            (agilereport_file, Path(self.agileReport_path).stem[:31]),
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
                                worksheet.set_column(col_num, col_num, max_len + 2) #Add padding

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
                                                              
                            
                            normalize = df is audit_results_wagers #Only normalize values for wager audit results
                            auditResults_versions_skipColumns = ['Jurisdiction', 'Game'] #Columns to specifically skip for audit_results_versions

                            #Iterates through rows/columns to apply formatting for mismatches
                            for row in range(1, len(df) + 1):
                                col_idx = 0 #Start at the first column
                                while col_idx < len(df.columns):
                                    try:
                                        remaining_columns = len(df.columns) - col_idx #Calculate remaining columns
                                        column_name = df.columns[col_idx]

                                        #Detect single columns for combined columns dynamically
                                        single_column = column_name in auditResults_versions_skipColumns or column_name == 'Game' or remaining_columns < 3
                                        #Excluding column name 'Jurisdiction' from being highlighted red due to inconsistencies on Agile PLM Report
                                        if single_column:
                                            val = df.iat[row - 1, col_idx]
                                            worksheet.write(row, col_idx, val)
                                            col_idx += 1
                                            continue

                                        #Access up to 3 values for comparison
                                        val1 = df.iat[row - 1, col_idx]
                                        val2 = df.iat[row - 1, col_idx + 1] if remaining_columns > 1 else None
                                        val3 = df.iat[row - 1, col_idx + 2] if remaining_columns > 2 else None

                                        if normalize: #Normalize values for audit_results_wagers only (if necessary)
                                            val1 = self.normalize_value(val1) if isinstance(val1, (int, float, str)) else val1
                                            val2 = self.normalize_value(val2) if isinstance(val2, (int, float, str)) else val2
                                            val3 = self.normalize_value(val3) if isinstance(val3, (int, float, str)) else val3

                                        #Replace NaN or None with empty string (if necessary)
                                        columns_in_groups = ["" if pd.isna(val1) or val1 is None else val1]
                                        if val2 is not None:
                                            columns_in_groups.append("" if pd.isna(val2) or val2 is None else val2)
                                        if val3 is not None:
                                            columns_in_groups.append("" if pd.isna(val3) or val3 is None else val3)

                                        #Only apply to audit_results_wagers and audit_results_versions
                                        if df is audit_results_wagers or df is audit_results_versions:
                                            n_vals = len(columns_in_groups)
                                            highlight_flags = [False] * n_vals #Flags for highlighting

                                            #Count occurrences of each value
                                            value_counts = {}
                                            for v in columns_in_groups:
                                                value_counts[v] = value_counts.get(v, 0) + 1

                                            max_frequency = max(value_counts.values()) #Find the majority value

                                            #Flag any value that is not in majority
                                            for i, v in enumerate(columns_in_groups):
                                                if value_counts[v] < max_frequency:
                                                    highlight_flags[i] = True

                                            #Write values to worksheet with red formatting
                                            for i, v in enumerate(columns_in_groups):
                                                fmt = red_format if highlight_flags[i] else None
                                                worksheet.write(row, col_idx + i, v, fmt)

                                            col_idx += n_vals #Move past this group

                                        else: #For all other sheets write normally without highlighting/comparison
                                            write_columns = min(3, remaining_columns)
                                            for i in range(write_columns):
                                                val = df.iat[row - 1, col_idx + i]
                                                worksheet.write(row, col_idx + i, val)
                                            col_idx += write_columns

                                    except Exception as e: #Debug to catch errors per row/column
                                        print(f"Error processing row: '{row}', col_idx: '{col_idx}': '{e}'")
                                        col_idx += 1 #Move to next column to avoid an infinite loop

                except Exception as e:
                    all_valid = False
                    messagebox.showerror("Error writing to Excel", str(e))
                    return False

            #Success message when results are True and all passes successfully
            if all_valid:
                messagebox.showinfo("Success!", "All files processed successfully and Wager & Game Version Audit Results are complete!")
                return True
            else:
                return False
            