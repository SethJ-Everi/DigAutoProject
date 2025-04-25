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
        self.root = root #Assigns Tkinter root window
        self.root.title("Wager Audit Comparison Tool") #Window title
        self.root.configure(bg="white") #Set window background color to white

        self.wageraudit_path = "" #path for Wager Audit csv file
        self.operatorsheet_path = "" #path for Op Wager Config Excel sheet file
        self.opgamelist_report_path = "" #path for Op GameList csv Report
        self.agilereport_path = "" #path for the Agile PLM Excel Report
        self.create_widgets() #method to create UI components

        #setting default and min size settings
        self.root.geometry("600x600") 
        self.root.minsize(600, 600)
        self.adjust_window() #method to center the window on the screen

    def adjust_window(self):
        #Get the screen's full width/height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 600
        window_height = 600

        #Calculate the top-left corner position to center the window
        position_top = int(screen_height / 2 - window_height / 2)
        position_left = int(screen_width / 2 - window_width / 2)

        #Update the window's geometry to apply size and position
        self.root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

    def create_widgets(self):
        #Welcome display text
        welcome_text = "\nWager Audit Comparison Tool"
 
        #Welcome label
        self.welcome_label = tk.Label(self.root, text=welcome_text, font=("TkDefaultFont", 12, "bold"), fg='black', bg='white')
        self.welcome_label.pack(pady=20)

        #Wager Audit label
        self.wageraudit_label = tk.Label(self.root, text="No csv file currently selected for the Wager Audit", fg='red', bg='white')
        self.wageraudit_label.pack(pady=2)

        #Wager Audit file upload button
        self.upload_wageraudit = tk.Button(self.root, text="Upload Wager Audit csv file", command=self.upload_wageraudit)
        self.upload_wageraudit.pack(pady=15)

        #Operator Wager Config Sheet label
        self.operatorsheet_label = tk.Label(self.root, text="No Excel file currently selected for the Operator Wager Config Sheet", fg='red', bg='white')
        self.operatorsheet_label.pack(pady=2)

        #Operator Sheet Button
        self.upload_operatorsheet = tk.Button(self.root, text="Upload Operator Wager Config Excel Sheet", command=self.upload_operatorsheet)
        self.upload_operatorsheet.pack(pady=15)

        #Operator GameList Report label
        self.opgamelist_report_label = tk.Label(self.root, text="No csv file currently selected for the Operator GameList Report", fg='red', bg='white')
        self.opgamelist_report_label.pack(pady=2)

        #Operator GameList Report button
        self.upload_opgamelist_report = tk.Button(self.root, text="Upload Operator GameList csv Report", command=self.upload_opgamelist_report)
        self.upload_opgamelist_report.pack(pady=15)

        #Agile PLM Report label
        self.agilereport_label = tk.Label(self.root, text="No Excel file currently selected for the Agile PLM Report", fg='red', bg='white')
        self.agilereport_label.pack(pady=2)

        #Agile PLM Report button
        self.upload_agilereport = tk.Button(self.root, text="Upload Agile PLM Excel Report", command=self.upload_agilereport)
        self.upload_agilereport.pack(pady=15)

        #Submit button
        self.submit_button = tk.Button(self.root, text="SUBMIT FILES", font=("TkDefaultFont", 12, "bold"), command=self.submit_files, state=tk.DISABLED)
        self.submit_button.pack(pady=30)

    def enable_submit_button(self):
        #Enables the submit button if both files are not empty and turns green. Displays red if only one is selected and remains disabled
        if self.wageraudit_path and self.operatorsheet_path and self.opgamelist_report_path and self.agilereport_path:
            self.submit_button.config(state=tk.NORMAL, bg='green', fg='black')
        else:
            self.submit_button.config(state=tk.DISABLED, bg='red', fg='black')

    def upload_wageraudit(self):
        #Allows user to upload excel or csv files
        self.wageraudit_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        #Checks if a file is selected
        if self.wageraudit_path:
            #Displays file name once selected and updates label from red to green
            self.wageraudit_label.config(text=f"Wager Audit csv file uploaded: \n{self.wageraudit_path.split('/')[-1]}", fg='green')
            
        else:
            #Show warning if no wager audit file is selected
            messagebox.showwarning("Missing Wager Audit csv file!", "Please select the Wager Audit csv file to proceed.")
            #Update label to indicate no file is selected and turn label text red
            self.wageraudit_label.config(text="No csv file currently selected for the Wager Audit.", fg='red')
            self.wageraudit_path = None

        #Enables submit button after selection
        self.enable_submit_button()

    def upload_operatorsheet(self):
        #Allows user to upload excel or csv files
        self.operatorsheet_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        #Checks if a file is selected
        if self.operatorsheet_path:
            #Displays file name once selected and updates label from red to green
            self.operatorsheet_label.config(text=f"Operator Wager Config Excel Sheet uploaded: \n{self.operatorsheet_path.split('/')[-1]}", fg='green')

        else:
            #Show warning if no operator wager config sheet is selected
            messagebox.showwarning("Missing Operator Wager Config Excel Sheet!", "Please select the Operator Wager Config Excel Sheet to proceed.")
            #Update label to indicate no file is selected and turn label text red
            self.operatorsheet_label.config(text="No Excel file currently selected for the Operator Wager Config Sheet.", fg='red')
            self.operatorsheet_path = None

        #Enables submit button after selection
        self.enable_submit_button()

    def upload_opgamelist_report(self):
        self.opgamelist_report_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        if self.opgamelist_report_path:
            self.opgamelist_report_label.config(text=f"Operator GameList csv Report uploaded: \n{self.opgamelist_report_path.split('/')[-1]}", fg='green')

        else:
            messagebox.showwarning("Missing Operator GameList Report!", "Please select the Operator GameList csv Report to proceed.")
            self.opgamelist_report_label.config(text="No csv file currently selected for the Operator GameList Report.", fg='red')
            self.opgamelist_report_path = None

        self.enable_submit_button()

    def upload_agilereport(self):
        self.agilereport_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        if self.agilereport_path:
            self.agilereport_label.config(text=f"Agile PLM Excel Report uploaded: \n{self.agilereport_path.split('/')[-1]}", fg='green')

        else:
            messagebox.showwarning("Missing Agile PLM Report!", "Please select the Agile PLM Excel Report to proceed.")
            self.agilereport_label.config(text="No Excel file currently selected for the Agile PLM Report.", fg='red')
            self.agilereport_path = None

        self.enable_submit_button()

    def submit_files(self):
        #Creates a tkinter root window
        root = tk.Tk()
        root.withdraw() #Hides the root window (so it doesn't pop up)

        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #file types filter
            title="Select location to save file" #dialog title
        )

        if not file_path:
            #If no file path selected, show cancelled message
            messagebox.showinfo("No file path selected!", "Select a file path and try again.")
            self.enable_submit_button()
            return

        #Message box to confirm files for submission and allows for user to hit cancel
        if messagebox.askyesno("Confirm", "Are you sure you would like to submit all the files selected for comparison?"):
            try:
                #Call the function to compare files and save
                result = self.compare_files(file_path)
                #Success message
                if result:
                    messagebox.showinfo("Saved!", f"File successfully saved at: {file_path}.")
                else:
                    #If the result indicates failure, show failure message
                    messagebox.showerror("Error!", "Failed to save file. Please fix errors and try again.")

            except Exception as e:
                #Show error if there's an exception while saving files
                messagebox.showerror("Error!", f"An error has occured during export: {str(e)}")
            
        else:
            #If user cancels, display cancelled message
            messagebox.showinfo("Cancelled!", "This has been cancelled. Please upload all required files and hit submit to try again.")

        #Resets submit button to it's default state after handling success, cancellation, or missing file path
        self.enable_submit_button()

    def normalize_name(self, name):
        #to standardize game name column; convert to lowercase, removes all spaces, removes apostrophes
        if isinstance(name, str):
            #Normalize any smart quotes or accents (Ex: Jack O'Lantern Jackpots)
            name = unicodedata.normalize('NFKD', name)
            #Remove spaces
            name = name.strip().replace(' ', '')
            #Remove straight and curly apostrophes using regex 
            name = re.sub(r"[’']", '', name)
            #Convert to lowercase
            return name.lower()
        return name
    
    #Standardize values to handle percentages, currencies, and NaN values
    def normalize_value(self, val):
        #Return empty string for NaN, empty string, or whitespace
        if pd.isna(val) or val == '' or val == ' ':
            return ''
        
        val = str(val).strip()
        
        #Handles converting percentages to decimals (ex: 90% -> 0.9)
        if isinstance(val, str) and '%' in val:
            try:
                decimal_val = float(val.replace('%', '').strip()) / 100 #convert to decimal
                return str(math.ceil(decimal_val * 100) / 100) #rounds up to the next decimal place (ex: 0.9595 to 0.96)
            except ValueError:
                return '' #if conversion fails, empty string
        
        #Handles multiple values separated by commas or space separated values (ex: $0.01, $0.05, $0.10, etc.)
        if any(sym in val for sym in ('$', '€', '£')):

            #Regex to detect multiple values vs single values
            currency_values = re.findall(r'[\$€£]?\d[\d,]*\.?\d*', val)          
            
            if len(currency_values) > 1:
                parts = [v.strip() for v in val.split(',')]
                normalized_values = [self.normalize_currency_values(p) for p in parts]
                return ','.join(normalized_values)
            else:
                return self.normalize_currency_values(val) #single value - normalize as one
        
        val = val.replace(' ', '')
        return self.clean_number_string(val)
    
    #Handles values without currency symbols such as default lines & bet multipliers
    def clean_number_string(self, val):
        try:
            num = float(val)
            if num.is_integer():
                return str(int(num))
            else:
                return str(num)
        except ValueError:
            return str(val).strip()

    #Helper method to handle currency symbols (ex: $€£) and commas
    def normalize_currency_values (self, val):
        try:
            #Remove the currency symbols/commas
            val = val.replace('$', '').replace('€', '').replace('£', '').replace(',', '').strip()
            num = float(val)
            if num.is_integer():
                return str(int(num))
            else:
                return "{:.2f}".format(num)

        except ValueError:
            return '' #if conversion fails, empty string
            
    #Handles automatically detecting header rows by scanning all rows
    def detect_header_row(self, file_path, header_indicator="Game"):

        #Read Excel file and strip leading/trailing whitespaces from string cells
        if file_path.endswith('.xlsx'):
            wager_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header
            wager_data = wager_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

        #Handles csv files differently
        elif file_path.endswith('.csv'):
            rows = [] #Empty list to store rows
            with open(file_path, 'r', encoding='ISO-8859-1') as f:
                #DEBUG to print first 5 lines from CSV/Wager Audit file:
                reader = csv.reader(f)
                print("\nDEBUG FOR WAGER AUDIT: Preview of raw CSV rows:")

                #Iterate over each row
                for i, row in enumerate(reader):
                    standardized_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]
                    
                    #DEBUG: print standardized row for first 5 rows
                    if i < 5:
                        print(f"Line {i}: {standardized_row}")
                    #Append normalized row to the list of rows
                    rows.append(standardized_row)
            #Convert rows to DataFrame after reading rows, replace empty strings, None values with NaN for easier handling
            wager_data = pd.DataFrame(rows).replace(['', None], np.nan)

        else:
            raise ValueError("Unsupported file format. Only csv and Excel files are supported.")
                        
        #Iterate through each row, convert all values to string, strip spaces
        for idx, row in wager_data.iterrows():
            row_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]

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
        if file_path.endswith('.xlsx'):
            version_data = pd.read_excel(file_path, header=None, engine='openpyxl') 
            version_data = version_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) 

        elif file_path.endswith('.csv'):
            rows = []
            with open(file_path, 'r', encoding='ISO-8859-1') as f:
                #DEBUG to print first 5 lines from CSV/VERSION REPORT:
                reader = csv.reader(f)
                print("\nDEBUG FOR GAME/MATH VERSIONS: Preview of raw CSV rows:")

                for i, row in enumerate(reader):
                    standardizedversion_row = [cell.strip() if isinstance(cell, str) and cell.strip() else '' for cell in row]
                    
                    #DEBUG: print standardized row for first 5 rows
                    if i < 5:
                        print(f"Line {i}: {standardizedversion_row}")

                    rows.append(standardizedversion_row)

            version_data = pd.DataFrame(rows).replace(['', None], np.nan)

        else:
            raise ValueError("Unsupported file format. Only csv and Excel files are supported.")
                        
        for idx, row in version_data.iterrows():
            versionrow_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            lowered_values = [val.lower() for val in versionrow_values]

            if any(header_version_indicator.lower() in val for val in lowered_values):
                print(f"Header row detected at index {idx}")
                return idx
            
        print("No matching header row found.")
        return None

    def compare_files(self, file_path):
            #Checks if Wager Audit csv File and the Operator Wager Config Excel Sheet have been uploaded
            if not self.wageraudit_path or not self.operatorsheet_path:
                messagebox.showerror("Error!", "Please upload both the Wager Audit csv File and the Operator Wager Config Excel Sheet to proceed.")
                return False
            #Checks if Operator GameList csv Report and the Agile PLM Excel Report have been uploaded
            if not self.opgamelist_report_path or not self.agilereport_path:
                messagebox.showerror("Error!", "Please upload both the Operator GameList csv Report and the Agile PLM Excel Report to proceed.")
                return False
            
            all_valid = True 
            
            #Step 1: process Wager Audit csv File and Operator GameList csv Report
            try:

                #Checks required columns are present in both files
                wageraudit_columns = ["Everi Game ID", "RTP MAX", "Denom", "Line Selection", "Bet Multiplier Selection", "Default Denom", "Default Line", 
                                      "Default Bet Multiplier", "Default Bet", "Min Bet", "Max Bet"]
            
                operatorsheet_columns = ["Game", "RTP%", "Denom Selection", "Line/Ways Selection", "Bet Multiplier Selection", "Default Denom Selection", "Default Line/Ways", 
                                        "Default Bet Multiplier", "Total Default Bet", "Min Bet", "Max Bet"]

                #Defining column mapping manually so that names match data
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

                #Detect the header rows for both files automatically finding column names
                wageraudit_header_row = self.detect_header_row(self.wageraudit_path, header_indicator="Everi Game ID")
                operatorsheet_header_row = self.detect_header_row(self.operatorsheet_path, header_indicator="Game")

                #Throws an error if no valid header rows are found in the files
                if wageraudit_header_row is None or operatorsheet_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows the Wager Audit csv file and the Operator Wager Config Excel Sheet.")
                    return False
            
                #Read full files, skipping the detected header rows
                if self.wageraudit_path.endswith('.xlsx'):
                    wageraudit_file = pd.read_excel(self.wageraudit_path, header=wageraudit_header_row, engine='openpyxl')
                elif self.wageraudit_path.endswith('.csv'):
                    wageraudit_file = pd.read_csv(self.wageraudit_path, skiprows=wageraudit_header_row, encoding='ISO-8859-1')
                else:
                    raise ValueError("Unsupported file format for Wager Audit csv file.")
                
                if self.operatorsheet_path.endswith('.xlsx'):
                    operatorsheet_file = pd.read_excel(self.operatorsheet_path, header=operatorsheet_header_row, engine='openpyxl')
                elif self.operatorsheet_path.endswith('.csv'):
                    operatorsheet_file = pd.read_csv(self.operatorsheet_path, header=operatorsheet_header_row, encoding='ISO-8859-1')
                else:
                    raise ValueError("Unsupported file format for Operator Wager Config Excel Sheet.")
                
                #Normalize column names, strip spaces
                wageraudit_file.columns = wageraudit_file.columns.astype(str).str.strip()
                operatorsheet_file.columns = operatorsheet_file.columns.astype(str).str.strip()

                #Filter only relevant columns
                wageraudit_file = wageraudit_file[wageraudit_columns]
                operatorsheet_file = operatorsheet_file[operatorsheet_columns]

                #Check for missing columns
                missing_wageraudit_columns = [col for col in wageraudit_columns if col not in wageraudit_file.columns]
                missing_operatorsheet_columns = [col for col in operatorsheet_columns if col not in operatorsheet_file.columns]

                #Checks for missing columns
                if missing_wageraudit_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Wager Audit csv file: {', '.join(missing_wageraudit_columns)}")
                    return False 
                if missing_operatorsheet_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Operator Wager Config Excel Sheet: {', '.join(missing_operatorsheet_columns)}")
                    return False
                              
                #Renames columns to match column mapping
                try:
                    wageraudit_file = wageraudit_file.rename(columns=column_mapping_wager)
                    operatorsheet_file = operatorsheet_file.rename(columns=column_mapping_wager)
                except Exception as e:
                    messagebox.showerror("Error in column_mapping_wager", str(e))
                    return False

                #Handles all missing columns by adding them with NaN values to both DataFrames
                for col in column_mapping_wager.values():
                    if col not in wageraudit_file.columns:
                        wageraudit_file[col] = pd.NA
                    if col not in operatorsheet_file.columns:
                        operatorsheet_file[col] = pd.NA

                #Applies normalization to columns
                wageraudit_file['Game'] = wageraudit_file['Game'].apply(self.normalize_name)
                operatorsheet_file['Game'] = operatorsheet_file['Game'].apply(self.normalize_name)
               
                #Fill NaN values with 'N/A' for consistency during comparison/export
                wageraudit_file = wageraudit_file.fillna('N/A')
                operatorsheet_file = operatorsheet_file.fillna('N/A')

                #Sorts Game columns alphabetically in all DataFrames
                wageraudit_file = wageraudit_file.sort_values(by='Game', ascending=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True)

                #Removes duplicates in both DataFrames to ensure it only appears once
                wageraudit_file = wageraudit_file.drop_duplicates(subset='Game')
                operatorsheet_file = operatorsheet_file.drop_duplicates(subset='Game')

                #Ensures DataFrames have only matching Game values
                common_games_wager = set(wageraudit_file['Game']).intersection(set(operatorsheet_file['Game']))

                #Find missing games from each sheet
                games_wageraudit_file = set(wageraudit_file['Game'])
                games_operatorsheet_file = set(operatorsheet_file['Game'])

                missing_games_wageraudit_file = games_wageraudit_file - games_operatorsheet_file
                missing_games_operatorsheet_file = games_operatorsheet_file - games_wageraudit_file

                #Build list of dicts to convert to DataFrame
                missing_games_wager = [{'Game': game, 'Status': 'Missing in Wager Audit csv file'} for game in missing_games_operatorsheet_file]
                missing_games_wager += [{'Game': game, 'Status': 'Missing in Operator Wager Config Excel file'} for game in missing_games_wageraudit_file]

                #Create DataFrame for sheet 5 (missing games for wager audit)
                missing_games_wager = pd.DataFrame(missing_games_wager).sort_values(by='Game').reset_index(drop=True)

                #Filer rows based on common games
                wageraudit_file = wageraudit_file[wageraudit_file['Game'].isin(common_games_wager)]
                operatorsheet_file = operatorsheet_file[operatorsheet_file['Game'].isin(common_games_wager)]

                #Sort both DataFrames by 'Game' column and reset index to maintain alignment
                wageraudit_file = wageraudit_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for sheet 3 to hold side-by-side columns for comparison
                audit_results_wagers = pd.DataFrame()

                #Single loop to handle renamed columns to normalize values and add columns side by side
                for wager_column in wageraudit_file.columns:
                    #Normalize wager audit file columns
                    wageraudit_file[wager_column] = wageraudit_file[wager_column].apply(self.normalize_value)

                    #Checks if the column exists in operatorsheet_file
                    if wager_column in operatorsheet_file.columns:
                        #Normalize operator wager config sheet columns
                        operatorsheet_file[wager_column] = operatorsheet_file[wager_column].apply(self.normalize_value)

                        #side by side columns from both sheets to the DataFrame
                        audit_results_wagers[f"{wager_column}\n(Wager Audit File): "] = wageraudit_file[wager_column]
                        audit_results_wagers[f"{wager_column}\n(Operator Wager Config Sheet): "] = operatorsheet_file[wager_column]
                    else:
                        print(f"'{wager_column}' not found in the Operator Wager Config Excel Sheet")

                        #Collect missing games from wager audit
                        missing_games_wager = pd.concat([missing_games_wager, pd.DataFrame({'Missing Games': [wager_column]})], ignore_index=True)

            except Exception as e:
                all_valid = False
                messagebox.showerror("Error", f"An error has occured for the Wager Audit csv file and the Operator Wager Config Excel Sheet: {str(e)}")
                return False
            
            #Step 2: process Operator GameList csv Report and Agile PLM Excel Report
            try:

                opgamelist_columns = ["jurisdictionId", "gameId", "mathVersion", "Version"]
                agilereport_columns = ["Jurisdiction", "GameName", "Math Version", "Latest Software Version"]

                column_mapping_versions = {
                    "jurisdictionId": "Jurisdiction",
                    "gameId": "GameName",
                    "mathVersion": "Math Version",
                    "Version": "Latest Software Version"
                }

                opgamelist_header_row = self.detect_version_row(self.opgamelist_report_path, header_version_indicator="jurisdictionId")
                agilereport_header_row = self.detect_version_row(self.agilereport_path, header_version_indicator="Jurisdiction")

                if opgamelist_header_row is None or agilereport_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows for the Operator GameList csv Report and the Agile PLM Excel Report.")
                    return False

                if self.opgamelist_report_path.endswith('.xlsx'):
                    opgamelist_file = pd.read_excel(self.opgamelist_report_path, header=opgamelist_header_row, engine='openpyxl', dtype=str)
                elif self.opgamelist_report_path.endswith('.csv'):
                    opgamelist_file = pd.read_csv(self.opgamelist_report_path, skiprows=opgamelist_header_row, encoding='ISO-8859-1', dtype=str)
                else:
                    raise ValueError("Unsupported file format for Operator GameList csv Report.")

                if self.agilereport_path.endswith('.xlsx'):
                    agilereport_file = pd.read_excel(self.agilereport_path, header=agilereport_header_row, engine='openpyxl', dtype=str)
                elif self.agilereport_path.endswith('.csv'):
                    agilereport_file = pd.read_csv(self.agilereport_path, header=agilereport_header_row, encoding='ISO-8859-1', dtype=str)
                else:
                    raise ValueError("Unsupported file format for Agile PLM Excel Report.")

                opgamelist_file.columns = opgamelist_file.columns.astype(str).str.strip()
                agilereport_file.columns = agilereport_file.columns.astype(str).str.strip()

                opgamelist_file = opgamelist_file[opgamelist_columns]
                agilereport_file = agilereport_file[agilereport_columns]

                try:
                    opgamelist_file = opgamelist_file.rename(columns=column_mapping_versions)
                    agilereport_file = agilereport_file.rename(columns=column_mapping_versions)         

                    if 'GameName' in opgamelist_file.columns:
                        opgamelist_file = opgamelist_file.rename(columns={'GameName': 'Game'})
                    if 'GameName' in agilereport_file.columns:
                        agilereport_file = agilereport_file.rename(columns={'GameName': 'Game'})

                except Exception as e:
                    messagebox.showerror("Error in column_mapping_versions", str(e))
                    return False
                
                final_column_names = list(opgamelist_file.columns)

                missing_opgamelist_columns = [col for col in final_column_names if col not in opgamelist_file.columns]
                missing_agilereport_columns = [col for col in final_column_names if col not in agilereport_file.columns]

                if missing_opgamelist_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Operator GameList csv Report: {', '.join(missing_opgamelist_columns)}")
                    return False
                
                if missing_agilereport_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Agile PLM Excel Report: {', '.join(missing_agilereport_columns)}")
                    return False

                for col in final_column_names:
                    if col not in opgamelist_file.columns:
                        opgamelist_file[col] = np.nan
                    if col not in agilereport_file.columns:
                        agilereport_file[col] = np.nan

                opgamelist_file['Game'] = opgamelist_file['Game'].apply(self.normalize_name)
                agilereport_file['Game'] = agilereport_file['Game'].apply(self.normalize_name)

                opgamelist_file = opgamelist_file.fillna('N/A')
                agilereport_file = agilereport_file.fillna('N/A')

                opgamelist_file = opgamelist_file.sort_values(by='Game', ascending=True)
                agilereport_file = agilereport_file.sort_values(by='Game', ascending=True)

                opgamelist_file = opgamelist_file.drop_duplicates(subset='Game')
                agilereport_file = agilereport_file.drop_duplicates(subset='Game')

                common_games_version = set(opgamelist_file['Game']).intersection(set(agilereport_file['Game']))

                games_opgamelist_file = set(opgamelist_file['Game'])
                games_agilereport_file = set(agilereport_file['Game'])

                missing_games_opgamelist_file = games_opgamelist_file - games_agilereport_file
                missing_games_agilereport_file = games_agilereport_file - games_opgamelist_file

                missing_games_versions = [{'Game': gameVersion, 'Status': 'Missing in Agile PLM Excel Report'} for gameVersion in missing_games_opgamelist_file]
                missing_games_versions += [{'Game': gameVersion, 'Status': 'Missing in Operator GameList csv Report'} for gameVersion in missing_games_agilereport_file]

                missing_games_versions = pd.DataFrame(missing_games_versions).sort_values(by='Game').reset_index(drop=True)

                opgamelist_file = opgamelist_file[opgamelist_file['Game'].isin(common_games_version)]
                agilereport_file = agilereport_file[agilereport_file['Game'].isin(common_games_version)]

                opgamelist_file = opgamelist_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                agilereport_file = agilereport_file.sort_values(by='Game', ascending=True).reset_index(drop=True)

                audit_results_versions = pd.DataFrame()

                for version_column in opgamelist_file.columns:
                    if version_column in agilereport_file.columns and version_column in agilereport_file.columns:

                        audit_results_versions[f"{version_column}\n(Operator GameList Report): "] = opgamelist_file[version_column]
                        audit_results_versions[f"{version_column}\n(Agile PLM Report): "] = agilereport_file[version_column]
                    else:
                        if version_column not in opgamelist_file.columns:
                            print(f"'{version_column}' not found in the Operator GameList Report")
                        if version_column not in agilereport_file.columns:
                            print(f"'{version_column}' not found in the Agile PLM Report")

                        #Collect missing games for version audit
                        missing_games_versions = pd.concat([missing_games_versions, pd.DataFrame({'Missing Games': [version_column]})], ignore_index=True)

            except Exception as e:
                all_valid = False
                messagebox.showerror("Error!", f"An error has occured for the Operator GameList csv Report and the Agile PLM Excel Report: {str(e)}")
                return False
            
            #Combine missing games from wager audit and versions audit
            combined_missing_games = pd.concat([missing_games_wager, missing_games_versions], ignore_index=True)
            
            #If all files are processed successfully, proceed with the Excel writing
            if all_valid:
                try:
                    #Write to excel with formatting
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        #Write to Excel with sheet names (based on selected file paths) truncated to 31 characters
                        wageraudit_file.to_excel(writer, sheet_name=Path(self.wageraudit_path).stem[:31], index=False) #wager audit csv file on sheet 1
                        operatorsheet_file.to_excel(writer, sheet_name=Path(self.operatorsheet_path).stem[:31], index=False) #op sheet config excel sheet on sheet 2
                        audit_results_wagers.to_excel(writer, sheet_name='Wager Audit Comparison Results', index=False) #Wager Audit Comparison Results with side by side comparison on sheet 3
                        opgamelist_file.to_excel(writer, sheet_name=Path(self.opgamelist_report_path).stem[:31], index=False) #op gamelist csv report on sheet 4
                        agilereport_file.to_excel(writer, sheet_name=Path(self.agilereport_path).stem[:31], index=False) #agile plm report on sheet 5
                        audit_results_versions.to_excel(writer, sheet_name='GameVersion Audit Results', index=False) #GameVersion Audit Results with side by side comparison on sheet 6
                        combined_missing_games.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games on sheet 7

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF0000'}) #Red format highlights cells red when there's a mismatch on the Wager Audit Comparison Results

                        #Loop & apply formats to all sheets
                        for df, sheet_name in [
                            (wageraudit_file, Path(self.wageraudit_path).stem[:31]),
                            (operatorsheet_file, Path(self.operatorsheet_path).stem[:31]),
                            (audit_results_wagers, 'Wager Audit Comparison Results'),
                            (opgamelist_file, Path(self.opgamelist_report_path).stem[:31]),
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

                            #Add filter to header row
                            worksheet.autofilter(0, 0, 0, len(df.columns) - 1)
                            #Freeze top row to keep headers visible when scrolling
                            worksheet.freeze_panes(1, 0)

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
                                    for col_idx in range(0, len(df.columns), 2): #Iterates over columns
                                        try:
                                            val1 = df.iloc[row - 1, col_idx]
                                            val2 = df.iloc[row - 1, col_idx + 1]

                                            #Excluding column name 'Jurisdiction' from being highlighted red due to inconsistencies on Agile PLM report
                                            column_name = df.columns[col_idx]
                                            if 'Jurisdiction' in column_name:
                                                continue

                                            if normalize and isinstance(val1, (int, float, str)) and isinstance(val2, (int, float, str)):
                                                #Normalize values if necessary
                                                val1 = self.normalize_value(val1)
                                                val2 = self.normalize_value(val2)

                                            #Convert NaN to empty strings for comparison
                                            safe_val1 = "" if pd.isna(val1) else val1
                                            safe_val2 = "" if pd.isna(val2) else val2

                                            #Checks for mismatches and highlights mismatches in red
                                            if safe_val1 != safe_val2:
                                                worksheet.write(row, col_idx, safe_val1, red_format)
                                                worksheet.write(row, col_idx + 1, safe_val2, red_format)
                                            else:
                                                worksheet.write(row, col_idx, safe_val1)
                                                worksheet.write(row, col_idx + 1, safe_val2)

                                        except Exception as e:
                                            print(f"Error processing row {row}, col_idx {col_idx}: {e}")

                except Exception as e:
                    all_valid = False
                    messagebox.showerror("Error writing to Excel", str(e))
                    return False

            #Success message when results are True and evrything passes successfully.
            if all_valid:
                messagebox.showinfo("Success!", "All files processed successfully and Final Audit Results are complete.")
                return True
            else:
                return False
            

if __name__ =="__main__":
    root = tk.Tk()
    app = CompareFiles(root)
    root.mainloop()