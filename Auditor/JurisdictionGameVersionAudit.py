import tkinter as tk #Tkinter library for building the GUI
from tkinter import filedialog, messagebox #File dialog and messagebox for interaction
import pandas as pd #Pandas library for handling Excel fles
import xlsxwriter #XlsxWriter library for writing Excel files with formatting
import re #Regular expression module used for pattern-based string manipulation
import unicodedata #Module for the Unicode Character Database
from pathlib import Path #Module for modern object-oriented way to handle filesystem paths
from difflib import SequenceMatcher #Import SequenceMatcher for computing similarity between two strings


class JurisdictionGameVersionAuditProgram:
    def __init__(self, master=None):
        self.window = tk.Toplevel(master)
        self.window.title("Jurisdiction Game Version Audit Comparison Tool") #Window title
        self.window.configure(bg="#2b2b2b") #Set window background color to white

        self.window.protocol("WM_DELETE_WINDOW", self.close_window) #X button will confirm if user wants to close

        self.supportPanel_report_path = "" #Path for installed game versions for all OPIDs retreived from Support Tool Admin Panel
        self.agileReport_path = "" #Path for the Agile PLM Report for Latest Software Versions
        self.create_widgets() #function for UI components
        self.adjust_window() #function for screen function

        #Default and min size settings
        self.window.geometry("600x500")
        self.window.minsize(600, 500)

    def close_window(self): #Function for cancel confirmation
        confirm = messagebox.askyesno(
            "Exit Jurisdiction Game Version Audit",
            "Are you sure you want to close the Jurisdiction Game Version Audit?"
        )
        if confirm:
            self.window.destroy() #To close this window only
        else:
            messagebox.showinfo(
                "Canceled!",
                "Close canceled."
            )
        
    def adjust_window(self):
        #Get the screens full width/height
        screen_width = self.window.winfo_screenwidth()
        screen_height = self.window.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 600
        window_height = 500

        #Calculate the top-left corner position to center the window
        position_top = int(screen_height / 2 - window_height / 2)
        position_left = int(screen_width / 2 - window_width / 2)

        #Update the windows geometry to apply size and position
        self.window.geometry(f'{screen_width}x{window_height}+{position_left}+{position_top}')

    def create_widgets(self):
        #Main content frame for all buttons/labels
        content_frame = tk.Frame(self.window, bg="#2b2b2b", height=300)
        content_frame.pack(fill="both", expand=True, padx=20, pady=10)

        #Welcome display text and label
        welcome_text = "\nJurisdiction Game Version\nAudit Comparison Tool\n"
        self.welcome_label = tk.Label(content_frame, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='white', bg='#2b2b2b')
        self.welcome_label.pack(pady=10)

        #Group container 
        group_container = tk.Frame(content_frame, bg="#2b2b2b")
        group_container.pack()

        #Center group for support panel/agile plm report buttons
        center_group = tk.LabelFrame(group_container, text="Jurisdiction Game Version Audit Files", font=("TkDefaultFont", 8, "bold"), fg='white', bd=3, relief="groove", bg="#2b2b2b", padx=10, pady=10)
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

        #Support Panel File label and upload button
        self.supportPanel_report_label = tk.Label(center_group, text="Select Support Panel File", **label_style)
        self.supportPanel_report_label.pack(pady=(0, 5))
        self.supportPanel_report_button = tk.Button(center_group, text="Upload Support Panel File", width=30, command=self.upload_supportPanel_report, **button_style)
        self.supportPanel_report_button.pack(pady=(0, 10))
        self.button_hover_effect(self.supportPanel_report_button)

        #Agile PLM Report label and upload button
        self.agileReport_label = tk.Label(center_group, text="Select Agile PLM Report", **label_style)
        self.agileReport_label.pack(pady=(0, 5))
        self.agileReport_button = tk.Button(center_group, text="Upload Agile PLM Report", width=30, command=self.upload_agileReport, **button_style)
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
        #Enables the submit button if both files are not empty and turns green
        if self.supportPanel_report_path and self.agileReport_path:
            self.submit_button.config(state=tk.NORMAL, bg='green')
        else:
            self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F") #Displays red if only one is selected and remains disabled

    def upload_supportPanel_report(self):
        self.supportPanel_report_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file only (this is the file type when file is downloadeded from admin panel)

        if self.supportPanel_report_path: #Checks if a file is selected
            self.supportPanel_report_label.config(text=f"Support Panel File Uploaded: \n{self.supportPanel_report_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Support Panel File to proceed.") #Show warning if no support panel file is selected 
            self.supportPanel_report_label.config(text="Select Support Panel File", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.supportPanel_report_path = "" if not self.supportPanel_report_path else self.supportPanel_report_path

        self.enable_submit_button() #Enables submit button after selection

    def upload_agileReport(self):
        self.agileReport_path = filedialog.askopenfilename(
            parent=self.window,
            filetypes=[("Excel Files", "*.xlsx")]
            ) #Allows user to upload excel file (this is the file type when file is downloadeded from agile power bi)

        if self.agileReport_path: #Checks if a file is selected
            self.agileReport_label.config(text=f"Agile PLM Report Uploaded: \n{self.agileReport_path.split('/')[-1]}", fg='#90EE90') #Displays file name once selected/updates label from red to green
        else:
            messagebox.showwarning("Missing File!", "Select Agile PLM Report to proceed.") #Show warning if no Agile PLM Report is selected
            self.agileReport_label.config(text="Select Agile PLM Report", fg='#FF6F6F') #Update label to indicate no file is selected/turn label text red
            self.agileReport_path = "" if not self.agileReport_path else self.agileReport_path
        self.enable_submit_button() #Enables submit button after selection

    def submit_files(self):
        #Checks if all files are uploaded
        if not all([self.supportPanel_report_path, self.agileReport_path]):
            messagebox.showwarning("Incomplete files!",
                                   "Upload all required files before submitting.") #Show warning if not all files were uploaded
            return
        
        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            parent=self.window,
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #File types filter
            title="File Save Location" #Dialog title
        )

        if not file_path:
            messagebox.showinfo("Missing File Path!",
                                "Select file path to save Jurisdiction Game Version Audit Results and try again.") #Show cancelled message if no save file path was selected
            self.enable_submit_button() #Enables submit button
            return

        #Message box to confirm user selected files for submission and allows user to hit cancel if needed to re-upload files
        if messagebox.askyesno("Confirm Submit",
                               "Are you sure you want to submit files for comparison?"):
            try:
                result = self.compare_files(file_path) #Call the function to compare files and save
                if result:
                    messagebox.showinfo("Jurisdiction Game Version Audit Results Saved!",
                                        f"Jurisdiction Game Version Audit Results file successfully saved at: {file_path}.") #Success message and show user save location
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
            #Clear all file paths
            self.supportPanel_report_path = ""
            self.agileReport_path = ""

            #Clear all labels and display red text
            self.supportPanel_report_label.config(text="Select Support Panel File", fg="#FF6F6F")
            self.agileReport_label.config(text="Select Agile PLM Report", fg="#FF6F6F")

            #Disable the submit button and turn red
            self.submit_button.config(state=tk.DISABLED, bg="#FF6F6F")

            #Show message box to user stating cleared files
            messagebox.showinfo("Cleared",
                                "All Files cleared. Select new files to upload.")
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
            
    def detect_header_row(self, file_path, header_version_indicator=["game_name", "GameName"]):
        #Handles automatically detecting header rows by scanning all rows and searching for header indicator
        if not file_path.endswith('.xlsx'):
            raise ValueError("Unsupported file format. Only Excel (.xlsx) files are supported.") #Raise error for incorrect file formats
        
        version_data = pd.read_excel(file_path, header=None, engine='openpyxl') #Checks all rows for header

        version_data = version_data.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x)) #Cleans up unwanted spaces before further processing

        #DEBUG: Print first 5 rows for inspection
        print("\nDEBUG: Preview of first 5 raw rows:")
        for i in range(min(5, len(version_data))):
            print(f"Row {i}: {version_data.iloc[i].tolist()}")
                        
        for idx, row in version_data.iterrows(): #Iterate through each row, convert all values to string, strip spaces
            versionrow_values = [str(cell).strip() for cell in row.values if isinstance(cell, str)]
            lowered_values = [val.lower() for val in versionrow_values]

            if any(header.lower() in val for header in header_version_indicator for val in lowered_values):
                print(f"Header row detected at index {idx}")
                return idx

        print("No matching header row found.")
        return None    

    def partialMatching_GameNames(self, supportPanel_report, agileReport, min_length_ratio=0.4):
        shorter, longer = sorted([supportPanel_report, agileReport], key=len) #Sort game names by length so 'shorter' is always the smaller one
        return shorter in longer and len(shorter) / len(longer) >= min_length_ratio #Checks for 1.substring match / 2.at least min length ratio of 50%

    def matching_GameNames(self, supportPanel_report_gameNames, agileReport_gameNames, threshold=85):
        #Handles Game Name matches that may be different on both reports. For ex, Off The Hook; Good Ol Fishin Hole in the Agile Report vs Good Ol Fishin Hole in the Support Panel File
        gameName_matches = [] #List to store final matches
        gameName_matches_agileReport = set() #Tracks game names from Agile PLM Report that have already been matched (avoid duplicates)

        #Loop over each game name in Support Panel File
        for supportPanel_report in supportPanel_report_gameNames:
            best_score = 0 #Reset best score for each game name in Support Panel File
            best_match = None #Reset best match for each game name in Support Panel File
            
            #Compare game name from Support Panel File against Agile PLM Report
            for agileReport in agileReport_gameNames:
                if agileReport in gameName_matches_agileReport: #Skip if Agile PLM Report game name matched another in Support Pane File
                    continue

                #Compute character-level similarity ratio between game names
                score = SequenceMatcher(None, supportPanel_report, agileReport).ratio() * 100

                #If this is the best match score, store it
                if score > best_score:
                    best_score = score
                    best_match = agileReport

                #Override: apply if no strong score but game names are partially similar ex: el dorado vs el dorado the lost city
                if self.partialMatching_GameNames(supportPanel_report, agileReport) and best_score < threshold:
                    best_score = threshold + 1
                    best_match = agileReport

            #If the best match is above threshold, store the match after checking all game names from both files
            if best_score >= threshold and best_match:
                gameName_matches.append((supportPanel_report, best_match, best_score)) #Add the successful match
                gameName_matches_agileReport.add(best_match) #Mark this as matched so it won't be reused

        return gameName_matches #Return the full list of matched game names and their scores

    def compare_files(self, file_path):
            #Checks if both files have been uploaded
            if not self.supportPanel_report_path or not self.agileReport_path:
                messagebox.showerror("Error!", "Upload all files to proceed.") #Show error message if no files were uploaded
                return False
            
            all_valid = True #Set the validation flag to True if all files are present and proceed with processing
            
            try:   
                #Detect the header rows for both files automatically finding column name 'game_name' for Support Panel File and 'GameName' for Agile PLM Report
                supportPanel_report_header_row = self.detect_header_row(self.supportPanel_report_path, header_version_indicator=["game_name"])
                agilereport_header_row = self.detect_header_row(self.agileReport_path, header_version_indicator=["GameName"])

                #Throws an error if no valid header rows are found in the files
                if supportPanel_report_header_row is None or agilereport_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows in all selected files.")
                    return False
            
                #Read full files, skipping the detected header rows
                if self.supportPanel_report_path.endswith('.xlsx'):
                    supportPanel_report_file = pd.read_excel(self.supportPanel_report_path, header=supportPanel_report_header_row, engine='openpyxl', dtype=str) #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Support Panel File. Only ('.xlsx') file type is supported.") #Raise error if incorrect file type is selected

                if self.agileReport_path.endswith('.xlsx'):
                    agileReport_file = pd.read_excel(self.agileReport_path, header=agilereport_header_row, engine='openpyxl', dtype=str) #File format is downloaded as xlsx therefore will only support this file type
                else:
                    raise ValueError("Unsupported file format for Agile PLM Report. Only ('.xlsx') file type is supported.") #Raise error if incorrect file type is selected

                #Renames columns to match column mapping; renames 'game_name' column from the Support Panel File and 'GameName' column from the Agile PLM Report to 'Game Name' for consistency
                try:
                    supportPanel_report_file.rename(columns={'game_name': 'Game Name'}, inplace=True)
                    agileReport_file.rename(columns={'GameName': 'Game Name'}, inplace=True)
                except Exception as e:
                    messagebox.showerror("Error!", f"Column renaming failed: {str(e)}") #Show error if there's an issue renaming columns to 'Game Name'
                    return False
                
                #Drop rows containing unwanted text from Agile PLM report specifically OR Blank rows completely
                unwanted_keywords_agilePLMReport = ['applied filters:']
                agileReport_file = agileReport_file[~agileReport_file.apply(
                    lambda row: (
                        any(isinstance(cell, str) and any(kw in cell.lower() for kw in unwanted_keywords_agilePLMReport) for cell in row)
                        or all(cell == "" or pd.isna(cell) for cell in row)
                    ),
                    axis=1
                )].reset_index(drop=True)

                #Normalize column names and strip spaces
                supportPanel_report_file.columns = supportPanel_report_file.columns.astype(str).str.strip()
                agileReport_file.columns = agileReport_file.columns.astype(str).str.strip()

                #Applies normalization to 'Game Name' column for both files
                supportPanel_report_file['Game Name'] = supportPanel_report_file['Game Name'].apply(self.normalize_name)
                agileReport_file['Game Name'] = agileReport_file['Game Name'].apply(self.normalize_name)
               
                #Fill NaN values with 'N/A' for consistency during comparison/export
                supportPanel_report_file = supportPanel_report_file.fillna('N/A')
                agileReport_file = agileReport_file.fillna('N/A')

                agileReport_file = agileReport_file.drop_duplicates(subset='Game Name', keep='last') #Keeps last listed version as it is the latest approved per the Agile PLM Report specifically

                #Sorts 'Game Name' column alphabetically in both DataFrames
                supportPanel_report_file = supportPanel_report_file.sort_values(by='Game Name', ascending=True)
                agileReport_file = agileReport_file.sort_values(by='Game Name', ascending=True)

                #Get Game Name matches from all files
                matches = self.matching_GameNames(
                    list(supportPanel_report_file['Game Name']),
                    list(agileReport_file['Game Name']),
                    threshold=85
                )

                #Get sets of matched Game Names
                matched_supportPanel_report_file = set([m[0] for m in matches])
                matched_agileReport_file = set([m[1] for m in matches])

                #Get all Game Names from both files
                games_supportPanel_report_file = set(supportPanel_report_file['Game Name'])
                games_agileReport_file = set(agileReport_file['Game Name'])

                #Determine missing Game Names from both files
                missing_games_supportPanel_report = games_agileReport_file - matched_agileReport_file
                missing_games_agileReport = games_supportPanel_report_file - matched_supportPanel_report_file

                allmissing_gameNames = [] #Empty dict to collect missing Game Names

                #Loop through all Game Names to see which are missing
                for game in missing_games_supportPanel_report:
                    allmissing_gameNames.append({'Game Name': game, 'Status': 'Missing in Support Panel File'})

                for game in missing_games_agileReport:
                    allmissing_gameNames.append({'Game Name': game, 'Status': 'Missing in Agile PLM Report'})

                #Final DataFrame for missing Game Names for sheet 4
                missing_games = pd.DataFrame(allmissing_gameNames).sort_values(by='Game Name').reset_index(drop=True)

                #Mapping from both files using matching_GameNames
                agileReport_to_supportPanel_map = {match[1]: match[0] for match in matches}

                #Preserve original Game Name column for sheet 2
                agileReport_file['Game_Name_Mapped'] = agileReport_file['Game Name']

                #Map Agile PLM Report Game Names to match Support Panel Game Names, store in Game_Name_Mapped
                agileReport_file['Game_Name_Mapped'] = agileReport_file['Game Name'].apply(lambda t: agileReport_to_supportPanel_map.get(t, None))

                #Filter matched Game Names only
                agileReport_file_matched = agileReport_file[agileReport_file['Game_Name_Mapped'].notnull()]
                supportPanel_report_file_matched = supportPanel_report_file[supportPanel_report_file['Game Name'].isin(matched_supportPanel_report_file)]

                #Merge using temporary mapped game name column and rename
                agileReport_file_reduced = agileReport_file_matched[['Game_Name_Mapped', 'Latest Software Version']]
                
                #Merge files with reduced columns and add next to 'Game Name'
                audit_results_versions = supportPanel_report_file_matched.merge(agileReport_file_reduced, left_on='Game Name', right_on='Game_Name_Mapped', how='left')

                #Drop extra Game_Name_Mapped column on merged files
                audit_results_versions.drop(columns=['Game_Name_Mapped'], inplace=True)

                #Sort results by ascending order
                audit_results_versions = audit_results_versions.sort_values(by='Game Name', ascending=True).reset_index(drop=True)

                #Rearrange columns putting Latest Software Version next to Game Name and adding to the final results on sheet 3
                cols = list(audit_results_versions.columns)
                cols.remove('Latest Software Version')
                gameName_index = cols.index('Game Name')
                cols.insert(gameName_index + 1, 'Latest Software Version')
                audit_results_versions = audit_results_versions[cols]

                #Dropping 'Game_Name_Mapped' from Agile PLM Report as it is no longer needed after this
                agileReport_file.drop(columns=['Game_Name_Mapped'], inplace=True)

            except Exception as e:
                all_valid = False
                print(f"Error caught in except block: {e}")
                messagebox.showerror("Error!", f"An error has occured for Support Panel File and Agile PLM Report: {str(e)}")
                return False

            if all_valid:
                try:
                    #Write to excel with formatting
                    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                        #Write to Excel with sheet names (based on selected file paths) truncated to 31 characters
                        supportPanel_report_file.to_excel(writer, sheet_name=Path(self.supportPanel_report_path).stem[:31], index=False) #Support Panel File raw data in sheet 1
                        agileReport_file.to_excel(writer, sheet_name=Path(self.agileReport_path).stem[:31], index=False) #Agile PLM Report raw data in sheet 2
                        audit_results_versions.to_excel(writer, sheet_name='Game Version Audit Results', index=False) #Game Version Audit Results in sheet 3
                        missing_games.to_excel(writer, sheet_name='Missing Games', index=False) #Missing games that don't exist in both file in sheet 4

                        #Access the workbook and worksheet to apply formatting
                        workbook = writer.book

                        #Define formats for header, cell, and red highlighting for mistmatched cell values
                        header_format = workbook.add_format({'bg_color': '#D9D9D9', 'bold': True, 'border': 2, 'text_wrap': True}) #Grey header format (bold, thick borders)
                        cell_format = workbook.add_format({'border': 1, 'border_color': '#BFBFBF'}) #borders for data cells
                        red_format = workbook.add_format({'bg_color': '#FF6F6F'}) #Red format highlights cells red when there's a mismatch on Game Version Audit Results

                        #Loop to apply formats to all sheets
                        for df, sheet_name in [
                            (supportPanel_report_file, Path(self.supportPanel_report_path).stem[:31]),
                            (agileReport_file, Path(self.agileReport_path).stem[:31]),
                            (audit_results_versions, 'Game Version Audit Results'),
                            (missing_games, 'Missing Games')
                        ]:
                            worksheet = writer.sheets[sheet_name]

                            #Header row formatting and auto adjust column widths
                            for col_num, column_name in enumerate(df.columns):
                                worksheet.write(0, col_num, column_name, header_format)

                                #Auto adjust column width to fit contents by calculating optimal column widths based on header/data length
                                if df[column_name].notna().any():
                                    max_val_len = df[column_name].astype(str).map(len).max()
                                else:
                                    max_val_len = 0

                                max_len = max(max_val_len, len(column_name))
                                worksheet.set_column(col_num, col_num, max_len + 2) #Add padding

                            worksheet.autofilter(0, 0, 0, len(df.columns) - 1) #Add filter to header row for user to be able to filter results as needed
                            worksheet.freeze_panes(1, 0) #Freeze header top row to keep headers visible when scrolling

                            #Write all data cells w/border formatting
                            for row in range(1, len(df) + 1):
                                for col in range(len(df.columns)):
                                    val = df.iat[row - 1, col]
                                    if pd.isna(val) or val in [float('inf'), float('-inf')]:
                                        worksheet.write(row, col, "", cell_format)
                                    else:
                                        worksheet.write(row, col, val, cell_format)

                            if df is audit_results_versions:
                                lastest_softwareVersion_idx =  df.columns.get_loc('Latest Software Version') #Get the index for Latest Software Version dynamically
                                columnCompare_indices = [i for i in range(lastest_softwareVersion_idx + 1, len(df.columns))] #Get indices of all columns after Latest Software Version to compare against

                                #Loop through each row of the DataFrame starting at row 1
                                for row in range(1, len(df) + 1):
                                    for col_idx in columnCompare_indices: #Loop through each column that comes after Latest Software Version
                                        try:
                                            val_latest_softwareVersion = df.iloc[row - 1, lastest_softwareVersion_idx] #Get value in Latest Software Version column
                                            safe_val_latest_softwareVersion = "" if pd.isna(val_latest_softwareVersion) else str(val_latest_softwareVersion).strip() #Convert to string, strip whitespaces, or use NaN for empty strings

                                            val_operatorVersions = df.iloc[row - 1, col_idx] #Get operator versions column
                                            safe_val_operatorVersions = "" if pd.isna(val_operatorVersions) else str(val_operatorVersions).strip() #Convert to string, strip whitespaces, or use NaN for empty strings

                                            #Compare against Latest Software Version and highlights mismatches in red
                                            if safe_val_latest_softwareVersion != safe_val_operatorVersions:
                                                worksheet.write(row, col_idx, safe_val_operatorVersions, red_format)
                                            else: #Leave and write normally if they match
                                                worksheet.write(row, col_idx, safe_val_operatorVersions)

                                        except Exception as e:
                                            print(f"Error processing row {row}, col_idx {col_idx}: {e}")

                except Exception as e:
                    all_valid = False
                    messagebox.showerror("Error writing to Excel", str(e))
                    return False

            #Success message when results are True and all passes successfully
            if all_valid:
                messagebox.showinfo("Success!", "All files processed successfully and Jurisdiction Game Version Audit Results are complete!")
                return True
            else:
                return False
