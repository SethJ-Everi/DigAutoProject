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
        self.window.title("Wager & Game/Math Version Audit Comparison Tool") #window title
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
            "Exit Wager & Game/Math Version Audit?",
            "Are you sure you want to close the Wager & Game/Math Version Audit?",
            parent=self.window
        )
        if confirm:
            self.window.destroy() #To close this window only
        else:
            messagebox.showinfo(
                "Canceled!",
                "Exit canceled.",
                parent=self.window
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
        welcome_text = "\nWager & Game/Math Version\nAudit Comparison Tool\n"
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

        #Button style dictionary for exit button
        exit_button_style = {
            "borderwidth": 1,
            "highlightthickness": 0,
            "font": ("TkDefaultFont", 10, "bold")
        }

        #Staging Wager Audit label and upload button
        self.wagerAudit_Staging_label = tk.Label(left_group, text="Staging Wager Audit File Uploaded: \nNONE", **label_style)
        self.wagerAudit_Staging_label.pack(pady=(0, 5))
        self.wagerAudit_Staging_button = tk.Button(left_group, text="Upload Staging Wager Audit File", width=38, command=self.upload_wagerAudit_Staging, **button_style)
        self.wagerAudit_Staging_button.pack(pady=(0, 10))
        self.button_hover_effect(self.wagerAudit_Staging_button)

        #Production Wager Audit label and upload button
        self.wagerAudit_Production_label = tk.Label(left_group, text="Production Wager Audit File Uploaded: \nNONE", **label_style)
        self.wagerAudit_Production_label.pack(pady=(10, 5))
        self.wagerAudit_Production_button = tk.Button(left_group, text="Upload Production Wager Audit File", width=38, command=self.upload_wagerAudit_Production, **button_style)
        self.wagerAudit_Production_button.pack(pady=(0, 10))
        self.button_hover_effect(self.wagerAudit_Production_button)

        #Operator Wager Config Sheet label and upload button
        self.operator_wagerSheet_label = tk.Label(left_group, text="Operator Wager Configuration Sheet Uploaded: \nNONE", **label_style)
        self.operator_wagerSheet_label.pack(pady=(10, 5))
        self.operator_wagerSheet_button = tk.Button(left_group, text="Upload Operator Wager Configuration Sheet", width=38, command=self.upload_operatorWagerSheet, **button_style)
        self.operator_wagerSheet_button.pack(pady=(0, 10))
        self.button_hover_effect(self.operator_wagerSheet_button)

        #Staging Operator GameList Report label and upload button
        self.opGameList_stagingReport_label = tk.Label(right_group, text="Staging GameList Report Uploaded: \nNONE", **label_style)
        self.opGameList_stagingReport_label.pack(pady=(0, 5))
        self.opGameList_stagingReport_button = tk.Button(right_group, text="Upload Staging GameList Report", width=38, command=self.upload_opGameList_stagingReport, **button_style)
        self.opGameList_stagingReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.opGameList_stagingReport_button)

        #Production Operator GameList Report label and upload button
        self.opGameList_productionReport_label = tk.Label(right_group, text="Production GameList Report Uploaded: \nNONE", **label_style)
        self.opGameList_productionReport_label.pack(pady=(10, 5))
        self.opGameList_productionReport_button = tk.Button(right_group, text="Upload Production GameList Report", width=38, command=self.upload_opGameList_productionReport, **button_style)
        self.opGameList_productionReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.opGameList_productionReport_button)

        #Agile PLM Report label and upload button
        self.agileReport_label = tk.Label(right_group, text="Agile PLM Report Uploaded: \nNONE", **label_style)
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

        #Exit button
        self.exit_button = tk.Button(content_frame, text="EXIT", width=20, command=self.close_window, bg="#FF6F6F", fg='white', **exit_button_style)
        self.exit_button.pack(padx=10)
        self.button_hover_effect(self.exit_button, normal_bg="#FF6F6F")

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
            self.wagerAudit_Staging_label.config(text="Staging Wager Audit File Uploaded: \nNONE", fg="#FF6F6F")
            self.wagerAudit_Production_label.config(text="Production Wager Audit File Uploaded: \nNONE", fg="#FF6F6F")
            self.operator_wagerSheet_label.config(text="Operator Wager Configuration Sheet Uploaded: \nNONE", fg="#FF6F6F")
            self.opGameList_stagingReport_label.config(text="Staging GameList Report Uploaded: \nNONE", fg="#FF6F6F")
            self.opGameList_productionReport_label.config(text="Production GameList Report Uploaded: \nNONE", fg="#FF6F6F")
            self.agileReport_label.config(text="Agile PLM Report Uploaded: \nNONE", fg="#FF6F6F")

            #Disable the submit button and turn red
            self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

            #Show message box to user stating cleared files
            messagebox.showinfo("All Files Cleared!",
                                "Cleared all uploaded files. Select new files to upload.")
            
        else: #Show message box to user the clear was canceled
            messagebox.showinfo("Canceled!",
                                "Clear canceled and uploaded files remain as is.")

    def upload_wagerAudit_Staging(self):
        self.wagerAudit_Staging_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.wagerAudit_Staging_path: #Checks if a file is selected
            self.wagerAudit_Staging_label.config(text=f"Staging Wager Audit File Uploaded: \n{self.wagerAudit_Staging_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("File Upload Canceled!", "File upload canceled. Select Staging Wager Audit File to upload.") #Show warning if no staging wager audit file is selected
            self.wagerAudit_Staging_label.config(text="Staging Wager Audit File Uploaded: \nNONE", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
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
            messagebox.showwarning("File Upload Canceled!", "File upload canceled. Select Production Wager Audit File to upload.") #Show warning if no production wager audit file is selected
            self.wagerAudit_Production_label.config(text="Production Wager Audit File Uploaded: \nNONE", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
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
            messagebox.showwarning("File Upload Canceled!", "File upload canceled. Select Operator Wager Configuration Sheet to upload.") #Show warning if no op wager config sheet is selected
            self.operator_wagerSheet_label.config(text="Operator Wager Configuration Sheet Uploaded: \nNONE", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.operator_wagerSheet_path = "" if not self.operator_wagerSheet_path else self.operator_wagerSheet_path
        self.enable_submit_button() #Enables submit button after selection

    def upload_opGameList_stagingReport(self):
        self.opGameList_stagingReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opGameList_stagingReport_path: #Checks if a file is selected
            self.opGameList_stagingReport_label.config(text=f"Staging GameList Report Uploaded: \n{self.opGameList_stagingReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("File Upload Canceled!", "File upload canceled. Select Staging GameList Report to upload.") #Show warning if no staging op gamelist report is selected
            self.opGameList_stagingReport_label.config(text="Staging GameList Report Uploaded: \nNONE", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opGameList_stagingReport_path = "" if not self.opGameList_stagingReport_path else self.opGameList_stagingReport_path
        self.enable_submit_button() #Enables submit button after selection
    
    def upload_opGameList_productionReport(self):
        self.opGameList_productionReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opGameList_productionReport_path: #Checks if a file is selected
            self.opGameList_productionReport_label.config(text=f"Production GameList Report Uploaded: \n{self.opGameList_productionReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("File Upload Canceled!", "File upload canceled. Select Production GameList Report to upload.") #Show warning if no production op gamelist report is selected
            self.opGameList_productionReport_label.config(text="Production GameList Report Uploaded: \nNONE", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opGameList_productionReport_path = "" if not self.opGameList_productionReport_path else self.opGameList_productionReport_path

    def upload_agileReport(self):
        self.agileReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file (this is the file type when file is downloaded from agile power bi)

        if self.agileReport_path: #Checks if a file is selected
            self.agileReport_label.config(text=f"Agile PLM Report Uploaded: \n{self.agileReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("File Upload Canceled!", "File upload canceled. Select Agile PLM Report to upload.") #Show warning if no agile plm report is selected
            self.agileReport_label.config(text="Agile PLM Report Uploaded: \nNONE", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.agileReport_path = "" if not self.agileReport_path else self.agileReport_path
        self.enable_submit_button() #Enables submit button after selection

    def submit_files(self):
        #Checks if all files are uploaded
        if not all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path, self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
            messagebox.showwarning("Incomplete files!",
                                   "Upload all required files before submitting.") #Show warning if not all files were uploaded
            return

        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            parent=self.window,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #file types filter
            title="File Save Location" #dialog title
        )

        if not file_path:
            messagebox.showinfo("No File Path Selected!",
                                "Select file path to save Wager & Game/Math Version Audit Results and try again.") #Show canceled message if no save file path was selected
            self.enable_submit_button() #Enables submit button
            return

        #Message box to confirm user selected files for submission and allows user to hit cancel if needed to re-upload files
        if messagebox.askyesno("Confirm Submit?",
                               "Are you sure you want to submit files for comparison?"):
            try:
                result = self.compare_files(file_path) #Call the function to compare files and save
                if result:
                    messagebox.showinfo("Full Audit Results Saved!",
                                        f"Wager & Game/Math Version Audit Results successfully saved at: {file_path}.") #Success message and show user save location
                else:
                    messagebox.showerror("Error!",
                                         "Failed to save file. Check the correct file formats were submitted and try again.") #Show failure message if results fail
            except Exception as e:
                messagebox.showerror("Error!",
                                     f"Error occurred during export: {str(e)}") #Show error if there's an exception while saving files
        else:
            messagebox.showinfo("Canceled!",
                                "File submission canceled. Upload all required files to submit and try again.") #Display cancel message if user hits cancel

        self.enable_submit_button() #Resets submit button to it's default state after handling success, cancellation, or missing file path

    def normalize_name(self, name): #Standardize game name column
        #Convert NaN or missing values to empty string
        if pd.isna(name) or str(name).strip() == '':
            return ''
        #Ensures all values are strings
        if not isinstance(name, str):
            name = str(name)
        name = unicodedata.normalize('NFKD', name) #Normalize any smart quotes or accents (Ex: Jack O'Lantern Jackpots)
        name = re.sub(r'(?<!^)(?=[A-Z][a-z])', ' ', name) #Split only before capital letters followed by lowercase (to avoid splitting acronyms)
        name = name.replace('_', ' ') #Remove underscores and adds a space (specific for postfix games)
        name = re.sub(r"[’';:]", '', name) #Remove straight and curly apostrophes using regex
        name = re.sub(r'\s+', '', name).strip() #Replace multiple spaces with no space, then strip leading/trailing
        return name.lower() #Convert to lowercase

    def normalize_value(self, val, is_percent_column=False):
        #Standardize values to handle percentages, currencies, and NaN values
        if pd.isna(val) or val == '' or val == ' ': #Return empty string for NaN, empty string, or whitespace
            return ''
        val = str(val).strip()

        #Handles converting percentages first
        if '%' in val:
            try:                
                #Check if it contains a decimal place and normalize these values only (ex: 93.94%)
                if '.' in val:
                    number_part = val.replace('%', '').strip() #Strip percent symbol
                    decimal_val = float(number_part) / 100 #Decimal percentage, divide by 100
                    rounded_val = math.floor(decimal_val * 100 + 0.5) / 100 #If RTP is above 0.5% round up; if below 0.5% round down (ex: 90.50% = 91%; 90.40% = 90%)
                    percent_val = int(rounded_val * 100) #Convert decimnal back to whole percent
                    return f"{percent_val}%" #Add back %
                else:
                    return val.strip() #Return as is stripping whitespace if already a percent

            except ValueError:
                return '' #If conversion fails, return empty string

        #Handle numeric percentages only (ex: 90%) for excel sheets (Op wager config sheet) that are converted to numeric floats (ex: 0.90)
        if is_percent_column:
            try:
                numeric_val = float(val)
                if 0 < numeric_val < 1: #Only decimals between 0 and 1
                    percent_val = int(math.floor(numeric_val * 100 + 0.5)) #If RTP is above 0.5% round up; if below 0.5% round down (ex: 90.50% = 91%; 90.40% = 90%)
                elif 1 <= numeric_val <= 100:
                    percent_val = int(math.floor(numeric_val + 0.5))
                else:
                    return ''
                val = f"{percent_val}%"
                return val
            except ValueError:
                pass #Not a numeric value continue

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

        #Handle comma-separated when currency is not included (ex: .25,1.00,5.00 -> 0.25,1,5)
        if ',' in val:
            parts = val.split(',')
            normalized_parts = [self.clean_number_string(p) for p in parts]
            val = ','.join(normalized_parts)
        else:
            val = self.clean_number_string(val)

        val = val.replace(' ', '')
        if any(char.isalpha() for char in val): #Capitalize any letters in the final normalized value (ex: 243Ways -> 243WAYS)
            val = ''.join(char.upper() if char.isalpha() else char for char in val)

        return val

    def clean_number_string(self, val):
        #Handles values without currency symbols such as default lines & bet multipliers (or data entries for denoms such as .25, 1.00, 5.00 -> 0.25,1,5)
        try:
            if val.startswith("."):
                val = "0" + val
            num = float(val) #Convert to float
            if num.is_integer(): #Checks if float is an integer
                return str(int(num)) #If integer, return as a string
            else:
                return str(num).rstrip("0").rstrip(".")
        except ValueError:
            return str(val).strip() #If conversion fails (val is not a number), return original value as a stripped string

    def normalize_currency_values(self, val):
    #Helper method to handle currency symbols (ex: $€£) and commas; can expand currencies as needed
        try:
            val = re.sub(r'[$€£,]', '', val).strip() #Remove the currency symbols/commas using regex
            if val.startswith("."):
                val = "0" + val
            num = float(val) #Convert to float
            if num.is_integer(): #Checks if float is an integer
                return str(int(num)) #If integer, convert to an integer then to a string (removes the decimal point)
            else:
                return str(num).rstrip("0").rstrip(".") #If not an integer, return as string formatted with two decimal places
        except ValueError:
            return '' #if conversion fails, return empty string

    def detect_header_row(self, file_path, header_indicator="Game"):
    #Handles automatically detecting header rows by scanning all rows for Wager Files
        if file_path.endswith('.xlsx'): #Read Excel file
            wager_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            wager_data = wager_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

            #DEBUG: Print first 5 rows for inspection
            print("\nDEBUG Excel Files: Preview of first 5 raw rows:")
            print(wager_data.head())

        elif file_path.endswith('.csv'): #Handles csv files differently
            rows = [] #Empty list to store rows
            with open(file_path, 'r', encoding='ISO-8859-1') as f: #DEBUG to print first 5 lines from csv file:
                reader = csv.reader(f)
                print("\nDEBUG CSV Files: Preview of first 5 raw rows:")

                for i, row in enumerate(reader): #Iterate over each row
                    standardized_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]

                    if i < 5: #DEBUG: print standardized row for first 5 rows
                        print(f"Line {i}: {standardized_row}")
                    rows.append(standardized_row) #Append normalized row to the list of rows

            #Convert rows to DataFrame after reading rows, replace empty strings, None values with NaN for easier handling
            wager_data = pd.DataFrame(rows).replace(['', None], np.nan)
        else:
            raise ValueError("Unsupported file format. Only ('.csv') and ('.xlsx') file types are supported.") #Raise error for incorrect file formats

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

    def detect_version_row(self, file_path, header_version_indicator="Jurisdiction"):
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
                print("\nDEBUG CSV Files: Preview of first 5 raw rows:")

                for i, row in enumerate(reader): #Iterate over each row
                    standardizedversion_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]

                    if i < 5: #DEBUG: print standardized row for first 5 rows
                        print(f"Line {i}: {standardizedversion_row}")

                    rows.append(standardizedversion_row) #Append normalized row to the list of rows

            #Convert rows to DataFrame after reading rows, replace empty strings, None values with NaN for easier handling
            version_data = pd.DataFrame(rows).replace(['', None], np.nan)
        else:
            raise ValueError("Unsupported file format. Only ('.csv') and ('.xlsx') file types are supported.") #Raise error for incorrect file formats

        for idx, row in version_data.iterrows(): #Iterate through each row, convert all values to string, strip spaces
            versionrow_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            lowered_values = [val.lower() for val in versionrow_values]

            #Check if 'Jurisdiction' is a part of any column names in this row
            if any(header_version_indicator.lower() in val for val in lowered_values):
                print(f"Header row detected at index {idx}")
                return idx

        print("No matching header row found.")
        return None

    def partialMatching_GameNames(self, opGameList_Staging, opGameList_Production, min_length_ratio=0.4):
        #Checks if shorter game name is contained in longer title and passes min length ratio
        shorter, longer = sorted([opGameList_Staging, opGameList_Production], key=len) #Sort game names by length so 'shorter' is always the smaller one
        return shorter in longer and len(shorter) / len(longer) >= min_length_ratio #Checks for 1.substring match / 2.at least min length ratio of 40%

    def matching_GameNames(self, opGameList_StagingReport_gameNames, opGameList_ProductionReport_gameNames, agileReport_gameNames=None, threshold=85):
        #Handles Game Name exact + partial matches for all three files
        gameName_matches = [] #Store final game name matches
        gameName_matches_opGameList_Production = set() #Tracks which game name matches have already been matched in the opGameList Production file
        used_agileReport = set() if agileReport_gameNames else None #Tracks used titles in Agile PLM Report file

        #Local forced matches mapping
        #Keys: (Staging file, Production file), Values: Agile PLM Report to force
        #Entered lowercase & without spaces since values are normalized before hitting this function
        forced_gameName_matches = {
            ("mgmlions", "mgmlions"): "detroitlionsdeluxe",
            ("borgata", "borgata"): "borgata777respin",
            ("mgmjets", "mgmjets"): "newyorkjetsdeluxe",
            ("mgmsteelers", "mgmsteelers"): "pittsburghsteelersdeluxe",
            ("nflphiladelphiaeagles", "nflphiladelphiaeagles"): "philadelphiaeaglesjackpots",
            ("cashmachinematchthree", "cashmachinematchthree"): "cashmachinematch3",
            ("hoopdynastymatchthree", "hoopdynastymatchthree"): "hoopdynastymatch3",
            ("doubleblackdiamondmatchthree", "doubleblackdiamondmatchthree"): "doubleblackdiamondmatch3"
        }

        #Loop through game names in opGameList Staging Report
        for opGameList_StagingReport in opGameList_StagingReport_gameNames:
            best_score2 = 0
            best_match2 = None

            #Compare Staging game names against all Production game names
            for opGameList_ProductionReport in opGameList_ProductionReport_gameNames:
                #Skip if Production game names already matched
                if opGameList_ProductionReport in gameName_matches_opGameList_Production:
                    continue

                #Similarity score (0-100)
                score2 = SequenceMatcher(None, opGameList_StagingReport, opGameList_ProductionReport).ratio() * 100
                #Boost score if partial match is detected
                if self.partialMatching_GameNames(opGameList_StagingReport, opGameList_ProductionReport):
                    score2 = max(score2, threshold + 1)

                #Keep best-scoring game name matches from Production file
                if score2 > best_score2:
                    best_score2 = score2
                    best_match2 = opGameList_ProductionReport

            #Lock in Production game name match if threshold is met
            if best_score2 >= threshold and best_match2:
                gameName_matches_opGameList_Production.add(best_match2)

                #Check for forced match for Agile PLM Report
                forced_gameNameMatches_agileReport = None
                if agileReport_gameNames:
                    forced_gameNameMatches_agileReport = forced_gameName_matches.get((opGameList_StagingReport, best_match2))

                #Add tuple with Agile PLM Report = Staging report to ensure display game name matches Staging/Production
                if forced_gameNameMatches_agileReport:
                    used_agileReport.add(forced_gameNameMatches_agileReport)
                    gameName_matches.append(
                        (opGameList_StagingReport, best_match2, forced_gameNameMatches_agileReport, best_score2, threshold + 1) #Agile PLM Report game name renamed to Staging game name
                    )
                    continue #Skip to next if statement

                #If Agile PLM Report game name already exists, try to match too
                if agileReport_gameNames:
                    best_score3 = 0
                    best_match3 = None
                    for agileReport in agileReport_gameNames:
                        if agileReport in used_agileReport: #Skip game names from Agile PLM Report already used
                            continue

                        #Similarity score (0-100)
                        score3 = SequenceMatcher(None, opGameList_StagingReport, agileReport).ratio() * 100
                        #Boost score if partial match is detected
                        if self.partialMatching_GameNames(opGameList_StagingReport, agileReport):
                            score3 = max(score3, threshold + 1)

                        #Keep best-scoring game name matches from Agile PLM Report
                        if score3 > best_score3:
                            best_score3 = score3
                            best_match3 = agileReport

                    #Save triple match if Agile PLM Report file passes threshold
                    if best_score3 >= threshold and best_match3:
                        used_agileReport.add(best_match3)
                        gameName_matches.append((opGameList_StagingReport, best_match2, best_match3, best_score2, best_score3))
                else:
                    #Save pair match (for Staging and Production files only)
                    gameName_matches.append((opGameList_StagingReport, best_match2, best_score2))

        #Return all collected game name matches
        return gameName_matches

    def compare_files(self, file_path):
            #Checks if all required files are missing
            if not all([self.wagerAudit_Staging_path, self.wagerAudit_Production_path, self.operator_wagerSheet_path, self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
                messagebox.showerror("Error!", "Upload all required files to proceed.") #Show error if any files are missing
                return False #Stop further execution if files are incomplete

            all_valid = True #Set the validation flag to True if all files are present and proceed with processing

            #Step 1: process Wager Staging/Production Audit Files and Operator Wager Config Sheet
            try:
                #Checks required columns are present in all files
                wagerAudit_columns = ["Everi Game ID", "RTP MAX", "Denom", "Line Selection", "Bet Multiplier Selection", "Default Denom", "Default Line", 
                                      "Default Bet Multiplier", "Default Bet", "Min Bet", "Max Bet"]

                operatorSheet_columns = ["Game", "RTP%", "Denom Selection", "Line/Ways Selection", "Bet Multiplier Selection", "Default Denom Selection", "Default Line/Ways", 
                                        "Default Bet Multiplier", "Total Default Bet", "Min Bet", "Max Bet"]

                #Defining column mapping for wager audit manually so that names match data
                column_mapping_wagerAudit = {
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
                wagerAudit_Staging_header_row = self.detect_header_row(self.wagerAudit_Staging_path, header_indicator="Everi Game ID")
                wagerAudit_Production_header_row = self.detect_header_row(self.wagerAudit_Production_path, header_indicator="Everi Game ID")
                operatorSheet_header_row = self.detect_header_row(self.operator_wagerSheet_path, header_indicator="Game")

                #Throws an error if no valid header rows are found in the files
                if wagerAudit_Staging_header_row is None or wagerAudit_Production_header_row is None or operatorSheet_header_row is None:
                    messagebox.showerror("Missing Header Rows!", "Could not find valid header rows in the Staging Wager Audit File, Production Wager Audit File, and Operator Wager Configuration Sheet.")
                    return False

                #Read full files, skipping the detected header rows
                if self.wagerAudit_Staging_path.endswith('.csv'):
                    wagerAudit_StagingFile = pd.read_csv(self.wagerAudit_Staging_path, skiprows=wagerAudit_Staging_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Staging Wager Audit File. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected

                #Read full files, skipping the detected header rows
                if self.wagerAudit_Production_path.endswith('.csv'):
                    wagerAudit_ProductionFile = pd.read_csv(self.wagerAudit_Production_path, skiprows=wagerAudit_Production_header_row, encoding='ISO-8859-1') #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Production Wager Audit File. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected

                if self.operator_wagerSheet_path.endswith('.xlsx'):
                    operatorSheet_file = pd.read_excel(self.operator_wagerSheet_path, header=operatorSheet_header_row, engine='openpyxl') #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Operator Wager Configuration Sheet. Only ('.xlsx') file type is supported.") #Raise error if incorrect file type is selected

                #Normalize column names, strip spaces
                wagerAudit_StagingFile.columns = wagerAudit_StagingFile.columns.astype(str).str.strip()
                wagerAudit_ProductionFile.columns = wagerAudit_ProductionFile.columns.astype(str).str.strip()
                operatorSheet_file.columns = operatorSheet_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                wagerAudit_StagingFile = wagerAudit_StagingFile[wagerAudit_columns]
                wagerAudit_ProductionFile = wagerAudit_ProductionFile[wagerAudit_columns]
                operatorSheet_file = operatorSheet_file[operatorSheet_columns]

                #Identify if expected columns are missing
                missing_wagerAudit_Staging_columns = [col for col in wagerAudit_columns if col not in wagerAudit_StagingFile.columns]
                missing_wagerAudit_Production_columns = [col for col in wagerAudit_columns if col not in wagerAudit_ProductionFile.columns]
                missing_operatorSheet_columns = [col for col in operatorSheet_columns if col not in operatorSheet_file.columns]

                #Checks for missing columns
                if missing_wagerAudit_Staging_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Staging Wager Audit File: {', '.join(missing_wagerAudit_Staging_columns)}")
                    return False
                if missing_wagerAudit_Production_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Production Wager Audit File: {', '.join(missing_wagerAudit_Production_columns)}")
                    return False
                if missing_operatorSheet_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Operator Wager Configuration Sheet: {', '.join(missing_operatorSheet_columns)}")
                    return False

                #Renames columns to match column mapping
                try:
                    wagerAudit_StagingFile = wagerAudit_StagingFile.rename(columns=column_mapping_wagerAudit)
                    wagerAudit_ProductionFile = wagerAudit_ProductionFile.rename(columns=column_mapping_wagerAudit)
                    operatorSheet_file = operatorSheet_file.rename(columns=column_mapping_wagerAudit)
                except Exception as e:
                    messagebox.showerror("Error in column_mapping_wagerAudit", str(e))
                    return False

                #Handles all missing columns by adding them with NaN values to both DataFrames
                for col in column_mapping_wagerAudit.values():
                    if col not in wagerAudit_StagingFile.columns:
                        wagerAudit_StagingFile[col] = pd.NA
                    if col not in wagerAudit_ProductionFile.columns:
                        wagerAudit_ProductionFile[col] = pd.NA
                    if col not in operatorSheet_file.columns:
                        operatorSheet_file[col] = pd.NA

                #Applies normalization to columns
                wagerAudit_StagingFile['Game'] = wagerAudit_StagingFile['Game'].apply(self.normalize_name)
                wagerAudit_ProductionFile['Game'] = wagerAudit_ProductionFile['Game'].apply(self.normalize_name)
                operatorSheet_file['Game'] = operatorSheet_file['Game'].apply(self.normalize_name)

                operatorSheet_file = operatorSheet_file[operatorSheet_file['Game'] != ''] #For blank game entries; drops them so they don't appear in the Missing Games sheet

                #Handle RTP% column for operatorSheet_file specifically
                percent_column = ['RTP%']

                #Skip Game column and normalize values in the other columns and fill NaN values with 'N/A'
                for wager_column in wagerAudit_StagingFile.columns:
                    if wager_column != 'Game':
                        wagerAudit_StagingFile[wager_column] = wagerAudit_StagingFile[wager_column].replace('', pd.NA).fillna('N/A').apply(self.normalize_value)
                for wager_column in wagerAudit_ProductionFile.columns:
                    if wager_column != 'Game':
                        wagerAudit_ProductionFile[wager_column] = wagerAudit_ProductionFile[wager_column].replace('', pd.NA).fillna('N/A').apply(self.normalize_value)
                for wager_column in operatorSheet_file.columns:
                    if wager_column != 'Game':
                        operatorSheet_file[wager_column] = operatorSheet_file[wager_column].replace('', pd.NA).fillna('N/A').apply(lambda x: self.normalize_value(x, is_percent_column=(wager_column in percent_column)))

                #Sorts Game columns alphabetically in all DataFrames
                wagerAudit_StagingFile = wagerAudit_StagingFile.sort_values(by='Game', ascending=True)
                wagerAudit_ProductionFile = wagerAudit_ProductionFile.sort_values(by='Game', ascending=True)
                operatorSheet_file = operatorSheet_file.sort_values(by='Game', ascending=True)

                #Removes duplicates in DataFrames to ensure it only appears once
                wagerAudit_StagingFile = wagerAudit_StagingFile.drop_duplicates(subset='Game')
                wagerAudit_ProductionFile = wagerAudit_ProductionFile.drop_duplicates(subset='Game')
                operatorSheet_file = operatorSheet_file.drop_duplicates(subset='Game')

                #File labels for labeling on Missing Games sheet
                file_labels = ['Staging Wager Audit File',
                               'Production Wager Audit File',
                               'Operator Wager Configuration Sheet']

                #Get sets of Game Names from each file
                games_wagerAudit_StagingFile = set(wagerAudit_StagingFile['Game'])
                games_wagerAudit_ProductionFile = set(wagerAudit_ProductionFile['Game'])
                games_operatorSheet_file = set(operatorSheet_file['Game'])

                #Ensures DataFrames have only matching Game values
                common_games_wager = games_wagerAudit_StagingFile & games_wagerAudit_ProductionFile & games_operatorSheet_file

                #Union of all Game Names across all three files
                all_games = sorted(games_wagerAudit_StagingFile | games_wagerAudit_ProductionFile | games_operatorSheet_file)
                allFiles_set = [games_wagerAudit_StagingFile, games_wagerAudit_ProductionFile, games_operatorSheet_file]

                allMissing_games = [] #Empty list to collect missing Game Names

                for game in all_games:
                    missing_in = [
                        file_labels[i] for i, game_sets in enumerate(allFiles_set)
                        if game not in game_sets
                    ]
                    #Append one row per Game Name with combined missing info
                    if missing_in:
                        combined_status = ', '.join(missing_in)
                        allMissing_games.append({
                            'Game': game,
                            'Status': f'Missing in {combined_status}'
                        })

                #Convert missing Game Names list of dicts into a DataFrame and sort it for Missing Games Sheet
                missing_games_wager = pd.DataFrame(allMissing_games).sort_values(by='Game').reset_index(drop=True)

                #Filter rows based on common Game Names in all three files and sort by game
                wagerAudit_StagingFile_matchedGameNames = wagerAudit_StagingFile[wagerAudit_StagingFile['Game'].isin(common_games_wager)].sort_values(by='Game', ascending=True).reset_index(drop=True)
                wagerAudit_ProductionFile_matchedGameNames = wagerAudit_ProductionFile[wagerAudit_ProductionFile['Game'].isin(common_games_wager)].sort_values(by='Game', ascending=True).reset_index(drop=True)
                operatorSheet_file_matchedGameNames = operatorSheet_file[operatorSheet_file['Game'].isin(common_games_wager)].sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for Wager Audit Results to hold side-by-side columns for comparison
                audit_results_wagerAudit = pd.DataFrame({'Game': wagerAudit_StagingFile_matchedGameNames['Game'].values})

                #Single loop to handle renamed columns to normalize values and add columns side by side
                for wager_column in wagerAudit_StagingFile_matchedGameNames.columns:
                    if wager_column == 'Game':
                        continue
                    if wager_column not in wagerAudit_StagingFile_matchedGameNames.columns:
                        raise KeyError(f"'{wager_column}' not found in 'wagerAudit_StagingFile_matchedGameNames' matched rows dataset.")
                    if wager_column not in wagerAudit_ProductionFile_matchedGameNames.columns:
                        raise KeyError(f"'{wager_column}' not found in 'wagerAudit_ProductionFile_matchedGameNames' matched rows dataset.")
                    if wager_column not in operatorSheet_file_matchedGameNames.columns:
                        raise KeyError(f"'{wager_column}' not found in 'operatorSheet_file_matchedGameNames' matched rows dataset.")
                    
                    if wager_column in wagerAudit_StagingFile_matchedGameNames.columns and wager_column in wagerAudit_ProductionFile_matchedGameNames.columns and wager_column in operatorSheet_file_matchedGameNames.columns:
                        wagerAudit_StagingFile_matchedGameNames[wager_column] = wagerAudit_StagingFile_matchedGameNames[wager_column].apply(self.normalize_value).reset_index(drop=True)
                        wagerAudit_ProductionFile_matchedGameNames[wager_column] = wagerAudit_ProductionFile_matchedGameNames[wager_column].apply(self.normalize_value).reset_index(drop=True)
                        operatorSheet_file_matchedGameNames[wager_column] = operatorSheet_file_matchedGameNames[wager_column].apply(lambda x: self.normalize_value(x, is_percent_column=(wager_column in percent_column))).reset_index(drop=True)

                        #Side by side columns from all sheets to the DataFrame
                        audit_results_wagerAudit[f"{wager_column}\n(Staging Wager Audit File): "] = wagerAudit_StagingFile_matchedGameNames[wager_column]
                        audit_results_wagerAudit[f"{wager_column}\n(Production Wager Audit File): "] = wagerAudit_ProductionFile_matchedGameNames[wager_column]
                        audit_results_wagerAudit[f"{wager_column}\n({Path(self.operator_wagerSheet_path).stem[:31]}): "] = operatorSheet_file_matchedGameNames[wager_column]

                audit_results_wagerAudit['Game'] = wagerAudit_StagingFile_matchedGameNames['Game'].values
                cols = list(audit_results_wagerAudit.columns)
                cols.remove('Game')
                cols.insert(0, 'Game')
                audit_results_wagerAudit = audit_results_wagerAudit[cols]

                audit_results_wagerAudit = audit_results_wagerAudit.sort_values(by='Game', ascending=True).reset_index(drop=True)

            except Exception as e:
                all_valid = False
                print(f"Error caught in except block: {e}")
                messagebox.showerror("Error", f"An error has occured for the Staging Wager Audit File, Production Wager Audit File, and Operator Wager Configuration Sheet: {str(e)}")
                return False

            #Step 2: Process Staging Operator GameList Report/Production Operator GameList Report, and Agile PLM Report
            try:
                #Checks required columns are present in all files
                opGameList_columns = ["jurisdictionId", "gameId", "mathVersion", "Version"]
                agileReport_columns = ["Jurisdiction", "GameName", "Math Version", "Latest Software Version"]

                #Defining column mapping for game/math version audit manually so that names match data
                column_mapping_gameVersionAudit = {
                    "jurisdictionId": "Jurisdiction",
                    "gameId": "GameName",
                    "mathVersion": "Math Version",
                    "Version": "Latest Software Version"
                }

                #Detect the header rows for files automatically finding column names
                opGameList_Staging_header_row = self.detect_version_row(self.opGameList_stagingReport_path, header_version_indicator="jurisdictionId")
                opGameList_Production_header_row = self.detect_version_row(self.opGameList_productionReport_path, header_version_indicator="jurisdictionId")
                agileReport_header_row = self.detect_version_row(self.agileReport_path, header_version_indicator="Jurisdiction")

                #Throws an error if no valid header rows are found in files
                if opGameList_Staging_header_row is None or opGameList_Production_header_row is None or agileReport_header_row is None:
                    messagebox.showerror("Missing Header Rows!", "Could not find valid header rows for the Staging GameList Report, Production GameList Report, and Agile PLM Report.")
                    return False

                #Read full files, skipping the detected header rows
                if self.opGameList_stagingReport_path.endswith('.csv'):
                    opGameList_StagingFile = pd.read_csv(self.opGameList_stagingReport_path, skiprows=opGameList_Staging_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Staging GameList Report. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected

                if self.opGameList_productionReport_path.endswith('.csv'):
                    opGameList_ProductionFile = pd.read_csv(self.opGameList_productionReport_path, skiprows=opGameList_Production_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Production GameList Report. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected

                if self.agileReport_path.endswith('.xlsx'):
                    agileReport_file = pd.read_excel(self.agileReport_path, header=agileReport_header_row, engine='openpyxl', dtype=str) #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Agile PLM Report. Only ('.xlsx') file type is supported.") #Raise error if incorrect file type is selected
                
                #Drop rows containing unwanted text from Agile PLM report specifically OR Blank rows completely
                unwanted_keywords_agilePLMReport = ['applied filters:']
                agileReport_file = agileReport_file[~agileReport_file.apply(
                    lambda row: (
                        any(isinstance(cell, str) and any(kw in cell.lower() for kw in unwanted_keywords_agilePLMReport) for cell in row)
                        or all(cell == "" or pd.isna(cell) for cell in row)
                    ),
                    axis=1
                )].reset_index(drop=True)

                #Normalize column names, strip spaces
                opGameList_StagingFile.columns = opGameList_StagingFile.columns.astype(str).str.strip()
                opGameList_ProductionFile.columns = opGameList_ProductionFile.columns.astype(str).str.strip()
                agileReport_file.columns = agileReport_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                opGameList_StagingFile = opGameList_StagingFile[opGameList_columns]
                opGameList_ProductionFile = opGameList_ProductionFile[opGameList_columns]
                agileReport_file = agileReport_file[agileReport_columns]

                #Identify if expected columns are missing
                missing_opGameList_Staging_columns = [col for col in opGameList_columns if col not in opGameList_StagingFile.columns]
                missing_opGameList_Production_columns = [col for col in opGameList_columns if col not in opGameList_ProductionFile.columns]
                missing_agileReport_columns = [col for col in agileReport_columns if col not in agileReport_file.columns]

                #Checks for missing columns and if missing, program will not continue
                if missing_opGameList_Staging_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Staging GameList Report: {', '.join(missing_opGameList_Staging_columns)}")
                    return False

                if missing_opGameList_Production_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Production GameList Report: {', '.join(missing_opGameList_Production_columns)}")
                    return False

                if missing_agileReport_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Agile PLM Report: {', '.join(missing_agileReport_columns)}")
                    return False

                #Renames columns to match column mapping; renames 'GameName' column to 'Game' for consistency
                try:
                    opGameList_StagingFile = opGameList_StagingFile.rename(columns=column_mapping_gameVersionAudit)
                    opGameList_ProductionFile = opGameList_ProductionFile.rename(columns=column_mapping_gameVersionAudit)
                    agileReport_file = agileReport_file.rename(columns=column_mapping_gameVersionAudit)

                    if 'GameName' in opGameList_StagingFile.columns:
                        opGameList_StagingFile = opGameList_StagingFile.rename(columns={'GameName': 'Game'})
                    if 'GameName' in opGameList_ProductionFile.columns:
                        opGameList_ProductionFile = opGameList_ProductionFile.rename(columns={'GameName': 'Game'})
                    if 'GameName' in agileReport_file.columns:
                        agileReport_file = agileReport_file.rename(columns={'GameName': 'Game'})

                except Exception as e:
                    messagebox.showerror("Error in column_mapping_gameVersionAudit", str(e)) #Throws error if column mapping fails
                    return False

                #Applies normalization to columns
                opGameList_StagingFile['Game'] = opGameList_StagingFile['Game'].apply(self.normalize_name)
                opGameList_ProductionFile['Game'] = opGameList_ProductionFile['Game'].apply(self.normalize_name)
                agileReport_file['Game'] = agileReport_file['Game'].apply(self.normalize_name)

                #Fill NaN values with 'N/A' for consistency during comparison/export
                opGameList_StagingFile = opGameList_StagingFile.fillna('N/A')
                opGameList_ProductionFile = opGameList_ProductionFile.fillna('N/A')
                agileReport_file = agileReport_file.fillna('N/A')

                #Removes duplicates in DataFrames to ensure it only appears once
                opGameList_StagingFile = opGameList_StagingFile.drop_duplicates(subset='Game')
                opGameList_ProductionFile = opGameList_ProductionFile.drop_duplicates(subset='Game')
                agileReport_file = agileReport_file.drop_duplicates(subset='Game', keep='last') #keeps last listed version as it is the latest approved per the Agile PLM Report specifically

                #Sorts 'Game' column alphabetically in DataFrames
                opGameList_StagingFile = opGameList_StagingFile.sort_values(by='Game', ascending=True)
                opGameList_ProductionFile = opGameList_ProductionFile.sort_values(by='Game', ascending=True)
                agileReport_file = agileReport_file.sort_values(by='Game', ascending=True)

                #File labels for labeling on Missing Games sheet
                file_labels = ['Staging GameList Report',
                               'Production GameList Report',
                               'Agile PLM Report']

                #Get Game Name matches from all files
                gameName_matches_gameVersionAudit = self.matching_GameNames(
                    list(opGameList_StagingFile['Game']),
                    list(opGameList_ProductionFile['Game']),
                    list(agileReport_file['Game']),
                    threshold=85,
                )

                #Build map for agile plm report game name to op gamelist staging to handle game name partial matches
                agileReportFile_to_opGameListStagingFile_map = {m[2]: m[0] for m in gameName_matches_gameVersionAudit if m[2] != m[0]}

                #Pre-align agile plm report game names using the mapping above
                agileReport_file_aligned = agileReport_file.copy()
                agileReport_file_aligned['Game'] = agileReport_file_aligned['Game'].apply(
                    lambda t: agileReportFile_to_opGameListStagingFile_map.get(t, t)
                )

                #Build missing game name sets for detection
                opGameList_stagingFile_set = set(opGameList_StagingFile['Game'])
                opGameList_productionFile_set = set(opGameList_ProductionFile['Game'])
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

                for t1, t2, t3, *_ in gameName_matches_gameVersionAudit:
                    row1_staging_df = opGameList_StagingFile.loc[opGameList_StagingFile['Game'] == t1]
                    row2_production_df = opGameList_ProductionFile.loc[opGameList_ProductionFile['Game'] == t2]
                    row3_agileReport_df = agileReport_file.loc[agileReport_file['Game'] == t3]

                    #Skip if any row is missing
                    if row1_staging_df.empty or row2_production_df.empty or row3_agileReport_df.empty:
                        continue

                    row1_idx_staging = row1_staging_df.iloc[0].to_dict()
                    row2_idx_production = row2_production_df.iloc[0].to_dict()
                    row3_idx_agileReport = row3_agileReport_df.iloc[0].copy().to_dict()

                    if t3 != t1:
                        row3_idx_agileReport['Game'] = t1

                    opGameList_stagingFile_matchedGameNames.append(row1_idx_staging)
                    opGameList_productionFile_matchedGameNames.append(row2_idx_production)
                    agileReport_matchedGameNames.append(row3_idx_agileReport)
                    gameName_rows.append({'Game': t1})

                #Build results table
                row1_staging_df = pd.DataFrame(opGameList_stagingFile_matchedGameNames).reset_index(drop=True)
                row2_production_df = pd.DataFrame(opGameList_productionFile_matchedGameNames).reset_index(drop=True)
                row3_agileReport_df = pd.DataFrame(agileReport_matchedGameNames).reset_index(drop=True)
                gameName_rows_df = pd.DataFrame(gameName_rows).drop_duplicates().reset_index(drop=True)
                audit_results_gameVersions = gameName_rows_df.copy()

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
                        raise KeyError(f"{col} not found in 'opGameList_StagingFile' matched rows datasets")
                    if col not in row2_production_df:
                        raise KeyError(f"{col} not found in 'opGameList_ProductionFile' matched rows datasets")
                    if col not in row3_agileReport_df:
                        raise KeyError(f"{col} not found in 'agileReport_file' matched rows datasets")

                    audit_results_gameVersions[f"{col}\n(Staging GameList Report): "] = row1_staging_df[col].reset_index(drop=True)
                    audit_results_gameVersions[f"{col}\n(Production GameList Report): "] = row2_production_df[col].reset_index(drop=True)
                    audit_results_gameVersions[f"{col}\n(Agile PLM Report): "] = row3_agileReport_df[col].reset_index(drop=True)

                #Combine Jurisdiction column to only appear once (pulled from agile plm report column)
                if 'Jurisdiction' in row3_agileReport_df.columns:
                    audit_results_gameVersions['Jurisdiction'] = row3_agileReport_df['Jurisdiction'].reset_index(drop=True)
                    audit_results_gameVersions = audit_results_gameVersions.sort_values(by='Game', ascending=True).reset_index(drop=True)
                    cols = list(audit_results_gameVersions.columns) #Rearrange columns putting Jurisdiction before Game
                    cols.remove('Jurisdiction')
                    gameName_index = cols.index('Game')
                    cols.insert(gameName_index, 'Jurisdiction')
                    audit_results_gameVersions = audit_results_gameVersions[cols]
                else:
                    audit_results_gameVersions = audit_results_gameVersions.sort_values(by='Game', ascending=True).reset_index(drop=True)

            except Exception as e:
                all_valid = False
                messagebox.showerror("Error!", f"An error has occured for the Staging Operator GameList Report, Production Operator GameList Report, and Agile PLM Report: {str(e)}")
                return False

            #Combine missing games from wager audit and game/math version audit combined for Missing Games sheet
            combined_missing_games = pd.concat([missing_games_wager, missing_gameNames_versionAudit], ignore_index=True)

            #If all files are processed successfully and True, proceed with Excel writing
            if all_valid:
                #Safety check to ensure file names are not the same so that it does not overwrite sheets accidently
                sheet_names_wagerAuditGroup = [
                    Path(self.wagerAudit_Staging_path).stem[:31],
                    Path(self.wagerAudit_Production_path).stem[:31],
                    Path(self.operator_wagerSheet_path).stem[:31]
                ]
                #Check for duplicates in wagerAuditGroup
                if len(sheet_names_wagerAuditGroup) != len(set(sheet_names_wagerAuditGroup)):
                    messagebox.showerror(
                        "Duplicate File Names Detected!",
                        f'Duplicate file names detected for files: {sheet_names_wagerAuditGroup}.\n'
                        'Rename files to ensure unique sheet names and re-upload again.'
                    )
                    return #Stop execution until files are renamed properly

                #Safety check to ensure file names are not the same for Game/Math Version Audit Files so that it does not overwrite sheets accidently
                sheet_names_gameVersionAuditGroup = [
                    Path(self.opGameList_stagingReport_path).stem[:31],
                    Path(self.opGameList_productionReport_path).stem[:31],
                    Path(self.agileReport_path).stem[:31]
                ]
                #Check for duplicates in gameVersionAuditGroup
                if len(sheet_names_gameVersionAuditGroup) != len(set(sheet_names_gameVersionAuditGroup)):
                    messagebox.showerror(
                        "Duplicate File Names Detected!",
                        f'Duplicate file names detected for files: {sheet_names_gameVersionAuditGroup}.\n'
                        'Rename files to ensure unique sheet names and re-upload again.'
                    )
                    return #Stop execution until files are renamed properly

                try:
                    #Write to excel with formatting
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        wagerAudit_StagingFile.to_excel(writer, sheet_name=sheet_names_wagerAuditGroup[0], index=False) #Staging Wager Audit File raw data on sheet 1
                        wagerAudit_ProductionFile.to_excel(writer, sheet_name=sheet_names_wagerAuditGroup[1], index=False) #Production Wager Audit File raw data on sheet 2
                        operatorSheet_file.to_excel(writer, sheet_name=sheet_names_wagerAuditGroup[2], index=False) #Op Wager Config Sheet raw data on sheet 3
                        audit_results_wagerAudit.to_excel(writer, sheet_name='Wager Audit Results', index=False) #Wager Audit Results with side by side comparison on sheet 4
                        opGameList_StagingFile.to_excel(writer, sheet_name=sheet_names_gameVersionAuditGroup[0], index=False) #Staging Op GameList Report raw data on sheet 5
                        opGameList_ProductionFile.to_excel(writer, sheet_name=sheet_names_gameVersionAuditGroup[1], index=False) #Production Op GameList Report raw data on sheet 6
                        agileReport_file.to_excel(writer, sheet_name=sheet_names_gameVersionAuditGroup[2], index=False) #Agile PLM Report raw data on sheet 7
                        audit_results_gameVersions.to_excel(writer, sheet_name='Game&Math Version Audit Results', index=False) #GameVersion Audit Results with side by side comparison on sheet 8
                        combined_missing_games.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games from all files on sheet 9

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #Borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF6F6F'}) #Red format highlights cells red when there's a mismatch on final audit results sheets

                        #Loop & apply formats to all sheets
                        for df, sheet_name in [
                            (wagerAudit_StagingFile, Path(self.wagerAudit_Staging_path).stem[:31]),
                            (wagerAudit_ProductionFile, Path(self.wagerAudit_Production_path).stem[:31]),
                            (operatorSheet_file, Path(self.operator_wagerSheet_path).stem[:31]),
                            (audit_results_wagerAudit, 'Wager Audit Results'),
                            (opGameList_StagingFile, Path(self.opGameList_stagingReport_path).stem[:31]),
                            (opGameList_ProductionFile, Path(self.opGameList_productionReport_path).stem[:31]),
                            (agileReport_file, Path(self.agileReport_path).stem[:31]),
                            (audit_results_gameVersions, 'Game&Math Version Audit Results'),
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

                            #Only freeze top row and game name column for wager sheets
                            if sheet_name in [
                                Path(self.wagerAudit_Staging_path).stem[:31],
                                Path(self.wagerAudit_Production_path).stem[:31],
                                Path(self.operator_wagerSheet_path).stem[:31],
                                "Wager Audit Results"
                            ]:
                                worksheet.freeze_panes(1, 1)

                            #Only freeze top row and first two columns in game version sheets
                            elif sheet_name in [
                                Path(self.opGameList_stagingReport_path).stem[:31],
                                Path(self.opGameList_productionReport_path).stem[:31],
                                Path(self.agileReport_path).stem[:31],
                                "Game&Math Version Audit Results"
                            ]:
                                worksheet.freeze_panes(1, 2)

                            else: #Only freeze top row for Missing Games sheet
                                worksheet.freeze_panes(1, 0)

                            #Write all data cells w/border formatting
                            for row in range(1, len(df) + 1):
                                for col in range(len(df.columns)):
                                    val = df.iat[row - 1, col]
                                    if pd.isna(val) or val in [float('inf'), float('-inf')]:
                                        worksheet.write(row, col, "", cell_format)
                                    else:
                                        worksheet.write(row, col, val, cell_format)

                            auditResults_versions_skipColumns = ['Jurisdiction', 'Game'] #Columns to specifically skip for audit_results_gameVersions from red highlighting

                            #Iterates through rows/columns to apply formatting for mismatches
                            for row in range(1, len(df) + 1):
                                col_idx = 0 #Start at the first column
                                while col_idx < len(df.columns):
                                    try:
                                        remaining_columns = len(df.columns) - col_idx #Calculate remaining columns
                                        column_name = df.columns[col_idx]

                                        #Detect single columns for combined columns dynamically
                                        single_column = column_name in auditResults_versions_skipColumns or column_name == 'Game' or remaining_columns < 3
                                        #Excluding column name 'Jurisdiction'/'Game' from being highlighted since columns are combined into one
                                        if single_column:
                                            val = df.iat[row - 1, col_idx]
                                            worksheet.write(row, col_idx, val)
                                            col_idx += 1
                                            continue

                                        #Access up to 3 values for comparison
                                        val1 = df.iat[row - 1, col_idx]
                                        val2 = df.iat[row - 1, col_idx + 1] if remaining_columns > 1 else None
                                        val3 = df.iat[row - 1, col_idx + 2] if remaining_columns > 2 else None

                                        #Replace NaN or None with empty string (if necessary)
                                        columns_in_groups = ["" if pd.isna(val1) or val1 is None else val1]
                                        if val2 is not None:
                                            columns_in_groups.append("" if pd.isna(val2) or val2 is None else val2)
                                        if val3 is not None:
                                            columns_in_groups.append("" if pd.isna(val3) or val3 is None else val3)

                                        #Only apply red highlithgting to audit_results_wagerAudit and audit_results_gameVersions
                                        if df is audit_results_wagerAudit or df is audit_results_gameVersions:
                                            n_vals = len(columns_in_groups)
                                            highlight_flags = [False] * n_vals #Flags for highlighting

                                            #Count occurrences of each value
                                            value_counts = {}
                                            for v in columns_in_groups:
                                                value_counts[v] = value_counts.get(v, 0) + 1

                                            #Highlight all values if all do not match
                                            if len(value_counts) ==  n_vals:
                                                highlight_flags = [True] * n_vals
                                            else:
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
                messagebox.showinfo("Success!", "All files processed successfully and Wager & Game/Math Version Audit Results are complete!")
                return True
            else:
                return False
