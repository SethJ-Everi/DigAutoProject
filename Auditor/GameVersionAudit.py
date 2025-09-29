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


class GameVersionAuditProgram:
    def __init__(self, master=None):
        self.window = tk.Toplevel(master)
        self.window.title("Game Version Audit Comparison Tool") #window title
        self.window.configure(bg="#2b2b2b") #set window background color to white

        self.window.protocol("WM_DELETE_WINDOW", self.close_window) #X button will confirm if user wants to close

        self.opGameList_stagingReport_path = "" #path for Staging Op GameList Report
        self.opGameList_productionReport_path = "" #path for Production Op GameList Report
        self.agileReport_path = "" #path for Agile PLM Report

        self.create_widgets() #function for UI components
        self.adjust_window() #function for screen function

        #Default and min size settings
        self.window.geometry("800x600")
        self.window.minsize(800, 600)

    def close_window(self): #function for cancel confirmation
        confirm = messagebox.askyesno(
            "Exit Game Version Audit",
            "Are you sure you want to close the Game Version Audit?"
        )
        if confirm:
            self.window.destroy() #To close this window only
        else:
            messagebox.showinfo(
                "Canceled!",
                "Close canceled."
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

        #Welcome display text
        welcome_text = "\nGame Version\nAudit Comparison Tool\n"
        self.welcome_label = tk.Label(content_frame, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='white', bg='#2b2b2b')
        self.welcome_label.pack(pady=10)

        #Group container
        group_container = tk.Frame(content_frame, bg="#2b2b2b")
        group_container.pack()

        #Center group for Staging Op GameList Report/Production Op GameList Report/Agile PLM Report
        center_group = tk.LabelFrame(group_container, text="Game Version Audit Files", font=("TkDefaultFont", 8, "bold"), fg='white', bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
        center_group.pack(anchor='center', padx=10)

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

        #Staging Operator GameList Report label and upload button
        self.opGameList_stagingReport_label = tk.Label(center_group, text="Select Staging Operator GameList Report", **label_style)
        self.opGameList_stagingReport_label.pack(pady=(0, 5))
        self.opGameList_stagingReport_button = tk.Button(center_group, text="Upload Staging Operator GameList Report", width=38, command=self.upload_opGameList_stagingReport, **button_style)
        self.opGameList_stagingReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.opGameList_stagingReport_button)

        #Production Operator GameList Report label and upload button
        self.opGameList_productionReport_label = tk.Label(center_group, text="Select Production Operator GameList Report", **label_style)
        self.opGameList_productionReport_label.pack(pady=(10, 5))
        self.opGameList_productionReport_button = tk.Button(center_group, text="Upload Production Operator GameList Report", width=38, command=self.upload_opGameList_productionReport, **button_style)
        self.opGameList_productionReport_button.pack(pady=(0, 10))
        self.button_hover_effect(self.opGameList_productionReport_button)

        #Agile PLM Report label and upload button
        self.agilereport_label = tk.Label(center_group, text="Select Agile PLM Report", **label_style)
        self.agilereport_label.pack(pady=(10, 5))
        self.agileReport_button = tk.Button(center_group, text="Upload Agile PLM Report", width=38, command=self.upload_agileReport, **button_style)
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
        if all([self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
            self.submit_button.config(state=tk.NORMAL, bg='green')
        else:
            self.submit_button.config(state=tk.DISABLED, bg='#FF6F6F')
        self.button_hover_effect(self.submit_button)

    def upload_opGameList_stagingReport(self):
        self.opGameList_stagingReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("CSV Files", "*.csv")]
            ) #Allows user to upload csv file (this is the file type when file is downloaded from admin panel)

        if self.opGameList_stagingReport_path: #Checks if a file is selected
            self.opGameList_stagingReport_label.config(text=f"Staging Operator GameList Report Uploaded: \n{self.opGameList_stagingReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Staging Operator GameList Report to proceed.") #Show warning if no staging op gamelist report is selected 
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
            messagebox.showwarning("Missing File!", "Select Production Operator GameList Report to proceed.") #Show warning if no production op gamelist report is selected 
            self.opGameList_productionReport_label.config(text="Select Production Operator GameList Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.opGameList_productionReport_path = "" if not self.opGameList_productionReport_path else self.opGameList_productionReport_path

    def upload_agileReport(self):
        self.agileReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file (this is the file type when file is downloaded from agile power bi)

        if self.agileReport_path: #Checks if a file is selected
            self.agilereport_label.config(text=f"Agile PLM Report Uploaded: \n{self.agileReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Agile PLM Report to proceed.") #Show warning if no agile plm report is selected
            self.agilereport_label.config(text="Select Agile PLM Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.agileReport_path = "" if not self.agileReport_path else self.agileReport_path
        self.enable_submit_button() #Enables submit button after selection

    def submit_files(self):
        #Checks if all files are uploaded
        if not all([self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
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
            messagebox.showinfo("Missing File Path!",
                                "Select file path to save Game Version Audit Results and try again.") #Show cancelled message if no save file path was selected
            self.enable_submit_button() #Enables submit button
            return

        #Message box to confirm user selected files for submission and allows user to hit cancel if needed to re-upload files
        if messagebox.askyesno("Confirm Submit",
                               "Are you sure you want to submit files for comparison?"):
            try:
                result = self.compare_files(file_path) #Call the function to compare files and save
                if result:
                    messagebox.showinfo("Game Version Audit Results Saved!",
                                        f"Game Version Audit Results successfully saved at: {file_path}.") #Success message and show user save location
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

    def clear_button(self):
        answer = messagebox.askyesno(
            "Confirm Clear",
            "Are you sure you want to clear all files selected?"
        )
        if answer:
            #Clear all file paths if yes is selected
            self.opGameList_stagingReport_path = ""
            self.opGameList_productionReport_path = ""
            self.agileReport_path = ""
            
            #Clear all labels and display red text
            self.opGameList_stagingReport_label.config(text="Select Staging Operator GameList Report", fg="#FF6F6F")
            self.opGameList_productionReport_label.config(text="Select Production Operator GameList Report", fg="#FF6F6F")
            self.agilereport_label.config(text="Select Agile PLM Report", fg="#FF6F6F")

            #Disable the submit button and turn red
            self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

            #Show message box to user stating cleared files
            messagebox.showinfo("All Files Cleared!",
                                "All uploaded files were cleared. Select new files to upload.")
            
        else: #Show message box to user the clear was canceled
            messagebox.showinfo("Canceled!",
                                "Clear canceled and files remain as is.")
        #Disable the submit button and turn red
        self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

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
                
    def detect_version_row(self, file_path, header_version_indicator="Jurisdiction"):
        #Handles automatically detecting header rows by scanning all rows
        if file_path.endswith('.xlsx'): #Read Excel file
            version_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            version_data = version_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

            #DEBUG: Print first 5 rows for inspection
            print("\nDEBUG Excel Files: Preview of first 5 raw rows:")
            print(version_data.head())

        elif file_path.endswith('.csv'): #Handles csv files differently
            rows = [] #Empty list to store rows
            with open(file_path, 'r', encoding='ISO-8859-1') as f:
                reader = csv.reader(f)
                print("\nDEBUG CSV Files: Preview of first 5 raw rows:") #DEBUG: Print first 5 rows for inspection

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
        shorter, longer = sorted([opGameList_Staging, opGameList_Production], key=len) #Sort game names by length so 'shorter' is always the smaller one
        return shorter in longer and len(shorter) / len(longer) >= min_length_ratio #Checks for 1.substring match / 2.at least min length ratio of 50%

    def matching_GameNames(self, opGameList_StagingReport_gameNames, opGameList_ProductionReport_gameNames, agileReport_gameNames=None, threshold=85):
        #Handles Game Name exact + partial matches for all three files
        gameName_matches = []
        gameName_matches_opGameList_Production = set()
        used_agileReport = set() if agileReport_gameNames else None

        for opGameList_StagingReport in opGameList_StagingReport_gameNames:
            best_score2 = 0
            best_match2 = None

            for opGameList_ProductionReport in opGameList_ProductionReport_gameNames:
                if opGameList_ProductionReport in gameName_matches_opGameList_Production:
                    continue

                score2 = SequenceMatcher(None, opGameList_StagingReport, opGameList_ProductionReport).ratio() * 100
                if self.partialMatching_GameNames(opGameList_StagingReport, opGameList_ProductionReport):
                    score2 = max(score2, threshold + 1)

                if score2 > best_score2:
                    best_score2 = score2
                    best_match2 = opGameList_ProductionReport

            if best_score2 >= threshold and best_match2:
                gameName_matches_opGameList_Production.add(best_match2)

                if agileReport_gameNames:
                    best_score3 = 0
                    best_match3 = None
                    for agileReport in agileReport_gameNames:
                        if agileReport in used_agileReport:
                            continue

                        score3 = SequenceMatcher(None, opGameList_StagingReport, agileReport).ratio() * 100
                        if self.partialMatching_GameNames(opGameList_StagingReport, agileReport):
                            score3 = max(score3, threshold + 1)

                        if score3 > best_score3:
                            best_score3 = score3
                            best_match3 = agileReport

                    if best_score3 >= threshold and best_match3:
                        used_agileReport.add(best_match3)
                        gameName_matches.append((opGameList_StagingReport, best_match2, best_match3, best_score2, best_score3))
                else:
                    gameName_matches.append((opGameList_StagingReport, best_match2, best_score2))

        return gameName_matches

    def compare_files(self, file_path):
            #Checks if all required files are missing
            if not all([self.opGameList_stagingReport_path, self.opGameList_productionReport_path, self.agileReport_path]):
                messagebox.showerror("Error!", "Upload all files to proceed.") #Show error if any files are missing
                return False #Stop further execution if files are incomplete
            
            all_valid = True #Set the validation flag to True if all files are present and proceed with processing
                        
            #Process for Staging Operator GameList/Production Operator GameList Reports and Agile PLM Report
            try:
                #Checks required columns are present in all files
                opGameList_columns = ["jurisdictionId", "gameId", "mathVersion", "Version"]
                agileReport_columns = ["Jurisdiction", "GameName", "Math Version", "Latest Software Version"]

                #Defining column mapping for version audit manually so that names match data
                column_mapping_versions = {
                    "jurisdictionId": "Jurisdiction",
                    "gameId": "GameName",
                    "mathVersion": "Math Version",
                    "Version": "Latest Software Version"
                }

                #Detect the header rows for files automatically finding column names
                opGameList_Staging_header_row = self.detect_version_row(self.opGameList_stagingReport_path, header_version_indicator="jurisdictionId")
                opGameList_Production_header_row = self.detect_version_row(self.opGameList_productionReport_path, header_version_indicator="jurisdictionId")
                agilereport_header_row = self.detect_version_row(self.agileReport_path, header_version_indicator="Jurisdiction")

                #Throws an error if no valid header rows are found in files
                if opGameList_Staging_header_row is None or opGameList_Production_header_row is None or agilereport_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows for Staging Operator GameList Report, Production Operator GameList Report, and Agile PLM Report.")
                    return False

                #Read full files, skipping the detected header rows
                if self.opGameList_stagingReport_path.endswith('.csv'):
                    opGameList_StagingFile = pd.read_csv(self.opGameList_stagingReport_path, skiprows=opGameList_Staging_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Staging Operator GameList Report. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected
                
                if self.opGameList_productionReport_path.endswith('.csv'):
                    opGameList_ProductionFile = pd.read_csv(self.opGameList_productionReport_path, skiprows=opGameList_Production_header_row, encoding='ISO-8859-1', dtype=str) #File format is downloaded as csv therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Production Operator GameList Report. Only ('.csv') file type is supported.") #Raise error if incorrect file type is selected

                if self.agileReport_path.endswith('.xlsx'):
                    agileReport_file = pd.read_excel(self.agileReport_path, header=agilereport_header_row, engine='openpyxl', dtype=str) #File format is downloaded as xlsx therefore will only support this file type
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
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Staging Operator GameList Report: {', '.join(missing_opGameList_Staging_columns)}")
                    return False
                
                if missing_opGameList_Production_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Production Operator GameList Report: {', '.join(missing_opGameList_Production_columns)}")
                    return False
                
                if missing_agileReport_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from Agile PLM Report: {', '.join(missing_agileReport_columns)}")
                    return False

                #Renames columns to match column mapping; renames 'GameName' column to 'Game' for consistency
                try:
                    opGameList_StagingFile = opGameList_StagingFile.rename(columns=column_mapping_versions)
                    opGameList_ProductionFile = opGameList_ProductionFile.rename(columns=column_mapping_versions)
                    agileReport_file = agileReport_file.rename(columns=column_mapping_versions)

                    if 'GameName' in opGameList_StagingFile.columns:
                        opGameList_StagingFile = opGameList_StagingFile.rename(columns={'GameName': 'Game'})
                    if 'GameName' in opGameList_ProductionFile.columns:
                        opGameList_ProductionFile = opGameList_ProductionFile.rename(columns={'GameName': 'Game'})
                    if 'GameName' in agileReport_file.columns:
                        agileReport_file = agileReport_file.rename(columns={'GameName': 'Game'})

                except Exception as e:
                    messagebox.showerror("Error in column_mapping_versions", str(e))
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
                file_labels = ['Staging Operator GameList Report',
                               'Production Operator GameList Report',
                               'Agile PLM Report']

                #Get Game Name matches from all files
                gameName_matches_versionAudit = self.matching_GameNames(
                    list(opGameList_StagingFile['Game']),
                    list(opGameList_ProductionFile['Game']),
                    list(agileReport_file['Game']),
                    threshold=85,
                )

                #Build map for agile plm report game name to op gamelist staging (partial matches)
                agileReport_file_to_opGameList_stagingFile_map = {m[2]: m[0] for m in gameName_matches_versionAudit if m[2] != m[0]}

                #Pre-align agile plm report game names using the mapping above
                agileReport_file_aligned = agileReport_file.copy()
                agileReport_file_aligned['Game'] = agileReport_file_aligned['Game'].apply(
                    lambda t: agileReport_file_to_opGameList_stagingFile_map.get(t, t)
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

                for t1, t2, t3, *_ in gameName_matches_versionAudit:
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
                        raise KeyError(f"{col} not found in 'opGameList_StagingFile' matched rows datasets")
                    if col not in row2_production_df:
                        raise KeyError(f"{col} not found in 'opGameList_ProductionFile' matched rows datasets")
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

                missing_games_versions = pd.DataFrame(allMissing_gameNames_versionAudit).sort_values(by='Game').reset_index(drop=True) #For missing game sheet
                
            except Exception as e:
                all_valid = False
                messagebox.showerror("Error!", f"An error has occured for the Staging Operator GameList Report, Production Operator GameList Report, and Agile PLM Report: {str(e)}")
                return False
            
            #If all files are processed successfully and True, proceed with Excel writing
            if all_valid:
                try:
                    #Write to excel with formatting
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        #Write to Excel with sheet names (based on selected file paths) truncated to 31 characters
                        opGameList_StagingFile.to_excel(writer, sheet_name=Path(self.opGameList_stagingReport_path).stem[:31], index=False) #Staging Op GameList Report raw data on sheet 1
                        opGameList_ProductionFile.to_excel(writer, sheet_name=Path(self.opGameList_productionReport_path).stem[:31], index=False) #Production Op GameList Report raw data on sheet 2
                        agileReport_file.to_excel(writer, sheet_name=Path(self.agileReport_path).stem[:31], index=False) #Agile PLM Report raw data on sheet 3
                        audit_results_versions.to_excel(writer, sheet_name='GameVersion Audit Results', index=False) #GameVersion Audit Results with side by side comparison on sheet 4
                        missing_games_versions.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games on sheet 5

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #Borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF6F6F'}) #Red format highlights cells red when there's a mismatch on GameVersion Audit Results

                        #Loop & apply formats to all sheets
                        for df, sheet_name in [
                            (opGameList_StagingFile, Path(self.opGameList_stagingReport_path).stem[:31]),
                            (opGameList_ProductionFile, Path(self.opGameList_productionReport_path).stem[:31]),
                            (agileReport_file, Path(self.agileReport_path).stem[:31]),
                            (audit_results_versions, 'GameVersion Audit Results'),
                            (missing_games_versions, 'Missing Games')
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
                                                              
                            
                            auditResults_versions_skipColumns = ['Jurisdiction', 'Game'] #Columns to specifically skip for audit_results_versions

                            #Iterates through rows/columns to apply formatting for mismatches
                            for row in range(1, len(df) + 1):
                                col_idx = 0 #Start at the first column
                                while col_idx < len(df.columns):
                                    try:
                                        remaining_columns = len(df.columns) - col_idx #Calculate remaining columns
                                        column_name = df.columns[col_idx]

                                        #Detect single columns for combined columns dynamically
                                        single_column = column_name in auditResults_versions_skipColumns or remaining_columns < 3
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

                                        #Replace NaN or None with empty string (if necessary)
                                        columns_in_groups = ["" if pd.isna(val1) or val1 is None else val1]
                                        if val2 is not None:
                                            columns_in_groups.append("" if pd.isna(val2) or val2 is None else val2)
                                        if val3 is not None:
                                            columns_in_groups.append("" if pd.isna(val3) or val3 is None else val3)

                                        #Only apply red highlighting to audit_results_versions
                                        if df is audit_results_versions:
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
                messagebox.showinfo("Success!", "All files processed successfully and Game Version Audit Results are complete.")
                return True
            else:
                return False
            
