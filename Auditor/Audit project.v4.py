import tkinter as tk #Tkinter library for building the GUI
from tkinter import filedialog, messagebox #file dialog and messagebox for interaction
import pandas as pd #Pandas library for handling Excel fles
import xlsxwriter #XlsxWriter library for writing Excel files with formatting
import math #module to allow for mathematical functions and constants
import re #Regular expression module used for pattern-based string manipulation

#Defines a class to compare files using graphical interface
class CompareFiles:
    def __init__(self, root):
        self.root = root #Assigns Tkinter root window
        self.root.title("Wager Audit Comparison Tool") #Window title

        self.wageraudit_path = "" #path for first file
        self.operatorsheet_path = "" #path for second file
        self.create_widgets() #method to create UI components

        #setting default and min size settings
        self.root.geometry("500x500") 
        self.root.minsize(500, 500)
        self.adjust_window() #method to center the window on the screen

    def adjust_window(self):
        #Get the screen's full width/height
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        #Defines the desired window dimensions
        window_width = 500
        window_height = 500

        #Calculate the top-left corner position to center the window
        position_top = int(screen_height / 2 - window_height / 2)
        position_left = int(screen_width / 2 - window_width / 2)

        #Update the window's geometry to apply size and position
        self.root.geometry(f'{window_width}x{window_height}+{position_left}+{position_top}')

    def create_widgets(self):
        #Welcome display text
        welcome_text = "Welcome to the Wager Audit Comparison Tool! \n\n" \
        "Select a Wager Audit file and \n" \
        "the Operator Wager Configuration Sheet.\n"
 
        #Welcome label
        self.welcome_label = tk.Label(self.root, text=welcome_text, font=("TkDefaultFont", 15, "bold"), fg='blue')
        self.welcome_label.pack(pady=20)

        #Wager Audit label
        self.wageraudit_label = tk.Label(self.root, text="No file selected for the Wager Audit", fg='red')
        self.wageraudit_label.pack(pady=5)

        #Wager Audit file upload button
        self.upload_wageraudit = tk.Button(self.root, text="Upload Wager Audit file", command=self.upload_wageraudit)
        self.upload_wageraudit.pack(pady=20)

        #Operator Wager Config Sheet label
        self.operatorsheet_label = tk.Label(self.root, text="No file selected for the Operator Wager Configuration Sheet", fg='red')
        self.operatorsheet_label.pack(pady=5)

        #Operator Sheet Button
        self.upload_operatorsheet = tk.Button(self.root, text="Upload Operator Wager Configuration Sheet", command=self.upload_operatorsheet)
        self.upload_operatorsheet.pack(pady=20)

        #Submit button
        self.submit_button = tk.Button(self.root, text="Submit both files for comparison", font=("TkDefaultFont", 12, "bold"), command=self.submit_files, state=tk.DISABLED)
        self.submit_button.pack(pady=20)

    def enable_submit_button(self):
        #Enables the submit button if both files are not empty and turns green. Displays red if only one is selected and remains disabled
        if self.wageraudit_path and self.operatorsheet_path:
            self.submit_button.config(state=tk.NORMAL, bg='green')
        else:
            self.submit_button.config(state=tk.DISABLED, bg='red')

    def upload_wageraudit(self):
        #Allows user to upload excel or csv files
        self.wageraudit_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        #Checks if a file is selected
        if self.wageraudit_path:
            #Displays file name once selected and updates label from red to green
            self.wageraudit_label.config(text=f"Wager Audit File Uploaded: {self.wageraudit_path.split('/')[-1]}", fg='green')
            
        else:
            #Show warning if no wager audit file is selected
            messagebox.showwarning("Wager Audit File missing.", "Please select a Wager Audit File.")
            #Update label to indicate no file is selected and turn label text red
            self.wageraudit_label.config(text="No file selected for the Wager Audit.", fg='red')
            self.wageraudit_path = None

        #Enables submit button after selection
        self.enable_submit_button()

    def upload_operatorsheet(self):
        #Allows user to upload excel or csv files
        self.operatorsheet_path = filedialog.askopenfilename(filetypes=[("Excel Files","*.xlsx;*.xls"), ("CSV Files", "*.csv")])

        #Checks if a file is selected
        if self.operatorsheet_path:
            #Displays file name once selected and updates label from red to green
            self.operatorsheet_label.config(text=f"Operator Wager Config Sheet Uploaded: {self.operatorsheet_path.split('/')[-1]}", fg='green')

        else:
            #Show warning if no operator wager config sheet is selected
            messagebox.showwarning("Operator Wager Config Sheet missing.", "Please select a Operator Wager Config Sheet.")
            #Update label to indicate no file is selected and turn label text red
            self.operatorsheet_label.config(text="No file selected for the Operator Wager Config Sheet.", fg='red')
            self.operatorsheet_path = None

        #Enables submit button after selection
        self.enable_submit_button()

    def submit_files(self):
        #Creates a tkinter root window
        root = tk.Tk()
        root.withdraw() #Hides the root window (so it doesn't pop up)

        #Allows user to select the file save location
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")], #file types filter
            title="Select location to save the Final Audit Results file." #dialog title
        )

        if not file_path:
            #If no file path selected, show cancelled message
            messagebox.showinfo("No file path selected", "Select a file and try again.")
            self.enable_submit_button()
            return
        
        #DEBUG LINE PRINT STATEMENT#
        try:
            with open(self.wageraudit_path, 'r', encoding='ISO-8859-1') as f:
                lines = f.readlines()
                print("First 5 lines:")
                print(lines[:5])
        except Exception as e:
            print(f"Error reading file: {e}")

        #Message box to confirm files for submission and allows for user to hit cancel
        if messagebox.askyesno("Confirm", "Are you sure you would like to submit the two files selected?"):
            try:
                #Call the function to compare files and save
                result = self.compare_files(file_path)
                #Success message
                if result:
                    messagebox.showinfo("Success!", f"Files have been submitted and saved at {file_path}.")
                else:
                    #If the result indicates failure, show failure message
                    messagebox.showerror("Error!", "Failed to save the file. Please fix errors and try again.")

            except Exception as e:
                #Show error if there's an exception while saving files
                messagebox.showerror("Error", f"An error has occured during export: {str(e)}")
            
        else:
            #If user cancels, display cancelled message
            messagebox.showinfo("Cancelled!", "This has been cancelled. Please hit submit when files are ready for comparison.")

        #Resets submit button to it's default state after handling success, cancellation, or missing file path
        self.enable_submit_button()

    def normalize_name(self, name):
        #to standardize game name column; convert to lowercase, removes all spaces, removes apostrophes
        if isinstance(name, str):
            return name.strip().replace(' ', '').replace("'", '').lower()
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
            
    #Handles automatically detecting header rows by scanning first 20 rows
    def detect_header_row(self, file_path, expected_columns):
        if file_path.endswith('.xlsx'):
            wager_data = pd.read_excel(file_path, header=None, nrows=20, engine='openpyxl') 
        elif file_path.endswith('.csv'):
            wager_data = pd.read_csv(file_path, header=None, nrows=20, encoding='ISO-8859-1')
        else:
            raise ValueError("Unsupported file format. Only .csv and .xlsx are supported.")
        
        #DEBUG LINE PRINT STATEMENT WHEN ISSUES ARISE DETECTING ROW HEADERS#
        print("Wager Data from the file:")
        print(wager_data.head(20))
                
        for idx, row in wager_data.iterrows():
            print(f"Checking row {idx}: {row.values}") #ADDITIONAL DEBUG LINE#
            if all(expected_col in row.values for expected_col in expected_columns):
                print(f"Header row detected at index {idx}")
                return idx
        return None

    def compare_files(self, file_path):
            #Checks if both files have been uploaded
            if not self.wageraudit_path or not self.operatorsheet_path:
                messagebox.showerror("Error!", "Please upload both the Wager Audit File and the Operator Wager Config Sheet for comparison.")
                return False
            
            try:
                #Checks required columns are present in both files
                wageraudit_columns = ["Everi Game ID", "RTP MAX", "Denom", "Line Selection", "Bet Multiplier Selection", "Default Denom", "Default Line", 
                                      "Default Bet Multiplier", "Default Bet", "Min Bet", "Max Bet"]
            
                operatorsheet_columns = ["Game", "RTP%", "Denom Selection", "Line/Ways Selection", "Bet Multiplier Selection", "Default Denom Selection", "Default Line/Ways", 
                                        "Default Bet Multiplier", "Total Default Bet", "Min Bet", "Max Bet"]
                
                #Defining column mapping manually so that names match data
                column_mapping = {
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
                           
                #Detect the header rows for both files automatically finding column name 'Everi Game ID' for the Wager Audit File and 'Game' for the Op sheet
                wageraudit_header_row = self.detect_header_row(self.wageraudit_path, ["Everi Game ID"])
                operatorsheet_header_row = self.detect_header_row(self.operatorsheet_path, ["Game"])

                #Throws an error if no valid header rows are found in the files
                if wageraudit_header_row is None or operatorsheet_header_row is None:
                    messagebox.showerror("Error!", "Could not find valid header rows in one or both files.")
                    return False
            
                #Read full files, skipping the detected header rows
                if self.wageraudit_path.endswith('.xlsx'):
                    wageraudit_file = pd.read_excel(self.wageraudit_path, header=wageraudit_header_row, engine='openpyxl')
                elif self.wageraudit_path.endswith('.csv'):
                    wageraudit_file = pd.read_csv(self.wageraudit_path, header=wageraudit_header_row, encoding='ISO-8859-1')
                else:
                    raise ValueError("Unsupported file format for Wager Audit file. Only .csv and .xlsx are supported.")

                if self.operatorsheet_path.endswith('.xlsx'):
                    operatorsheet_file = pd.read_excel(self.operatorsheet_path, header=operatorsheet_header_row, engine='openpyxl')
                elif self.operatorsheet_path.endswith('.csv'):
                    operatorsheet_file = pd.read_csv(self.operatorsheet_path, header=operatorsheet_header_row, encoding='ISO-8859-1')
                else:
                    raise ValueError("Unsupported file format for Operator Wager Configuration Sheet. Only .csv and .xlsx are supported.")


                #Normalize column names and strip spaces
                wageraudit_file.columns = wageraudit_file.columns.astype(str).str.strip()
                operatorsheet_file.columns = operatorsheet_file.columns.astype(str).str.strip()

                #Filters DataFrames to include only relevant columns
                wageraudit_file = wageraudit_file[wageraudit_columns]
                operatorsheet_file = operatorsheet_file[operatorsheet_columns]

                #Check that all expected columns are present in both DataFrames
                missing_wageraudit_columns = [col for col in wageraudit_columns if col not in wageraudit_file.columns]
                missing_operatorsheet_columns = [col for col in operatorsheet_columns if col not in operatorsheet_file.columns]

                #Checks for missing columns
                if missing_wageraudit_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Wager Audit file: {', '.join(missing_wageraudit_columns)}")
                    return False 
                if missing_operatorsheet_columns:
                    messagebox.showerror("Missing columns!", f"The following columns are missing from the Operator Wager Config Sheet: {', '.join(missing_operatorsheet_columns)}")
                    return False

                #Renames columns to match column mapping
                wageraudit_file = wageraudit_file.rename(columns=column_mapping)
                operatorsheet_file = operatorsheet_file.rename(columns=column_mapping)

                #Handles all missing columns by adding them with NaN values to both DataFrames
                for col in column_mapping.values():
                    if col not in wageraudit_file.columns:
                        wageraudit_file[col] = pd.NA
                    if col not in operatorsheet_file.columns:
                        operatorsheet_file[col] = pd.NA

                #Applies normalization to 'Game' column for both files
                wageraudit_file['Game'] = wageraudit_file['Game'].apply(self.normalize_name)
                operatorsheet_file['Game'] = operatorsheet_file['Game'].apply(self.normalize_name)
               
                #Fill NaN values with 'N/A' for consistency during comparison/export
                wageraudit_file = wageraudit_file.fillna('N/A')
                operatorsheet_file = operatorsheet_file.fillna('N/A')

                #Sorts 'Game' column alphabetically in both DataFrames
                wageraudit_file = wageraudit_file.sort_values(by='Game', ascending=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True)

                #Removes duplicates in both DataFrames to ensure it only appears once
                wageraudit_file = wageraudit_file.drop_duplicates(subset='Game')
                operatorsheet_file = operatorsheet_file.drop_duplicates(subset='Game')

                #Ensures both DataFrames have only matching Game values
                common_games = set(wageraudit_file['Game']).intersection(set(operatorsheet_file['Game']))

                #Filer rows based on common games
                wageraudit_file = wageraudit_file[wageraudit_file['Game'].isin(common_games)]
                operatorsheet_file = operatorsheet_file[operatorsheet_file['Game'].isin(common_games)]

                #Sort both DataFrames by 'Game' column and reset index to maintain alignment
                wageraudit_file = wageraudit_file.sort_values(by='Game', ascending=True).reset_index(drop=True)
                operatorsheet_file = operatorsheet_file.sort_values(by='Game', ascending=True).reset_index(drop=True)

                #DataFrame for sheet 3 to hold side-by-side columns for comparison
                audit_results = pd.DataFrame()

                #Single loop to handle renamed columns to normalize values and add columns side by side
                for wager_column in wageraudit_file.columns:
                    #Normalize wager audit file columns
                    wageraudit_file[wager_column] = wageraudit_file[wager_column].apply(self.normalize_value)

                    #Checks if the column exists in operatorsheet_file
                    if wager_column in operatorsheet_file.columns:
                        #Normalize operator wager config sheet columns
                        operatorsheet_file[wager_column] = operatorsheet_file[wager_column].apply(self.normalize_value)

                        #side by side columns from both sheets to the DataFrame
                        audit_results[f"{wager_column} (Wager Audit File): "] = wageraudit_file[wager_column]
                        audit_results[f"{wager_column} (Operator Wager Config Sheet): "] = operatorsheet_file[wager_column]
                    else:
                        print(f"'{wager_column}' not found in the Operator Wager Config Sheet")

                #Write to excel with formatting
                with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                    #Write original files to their respective sheets
                    wageraudit_file.to_excel(writer, sheet_name='Wager Audit File', index=False)
                    operatorsheet_file.to_excel(writer, sheet_name='Operator Wager Config Sheet', index=False)
                    #Final audit results with side by side comparison
                    audit_results.to_excel(writer, sheet_name='Final Audit Results', index=False)

                    #Access the workbook and worksheet to apply formatting
                    workbook = writer.book
                    worksheet = writer.sheets['Final Audit Results']
                
                    #Red format which highlights cells red when there's a mismatch on the Final Audit Results
                    red_format = workbook.add_format({'bg_color': '#FF0000'})
              
                    #Iterates through rows/columns to apply formatting for mismatches
                    for row in range(1, len(audit_results) + 1):
                        for col_idx in range(0, len(audit_results.columns), 2): #Iterates over Wager vs Operator columns
                            try:
                                val1 = audit_results.iloc[row - 1, col_idx]
                                val2 = audit_results.iloc[row - 1, col_idx + 1]

                                if isinstance(val1, (int, float, str)) and isinstance(val2, (int, float, str)):
                                    #Normalize values if necessary
                                    val1 = self.normalize_value(val1)
                                    val2 = self.normalize_value(val2)
                                #Checks for mismatches and highlights mismatches in red
                                if val1 != val2:
                                    print(f"Mismatch found for row {row}, col_idx {col_idx}, val1 = {val1}, val2 = {val2}")
                                    worksheet.write(row, col_idx, val1, red_format)
                                    worksheet.write(row, col_idx + 1, val2, red_format)
                                else:
                                    print(f"Skipping non-iterable values at row {row}, col_idx {col_idx}, val1 = {val1}, val2 = {val2}")

                            except Exception as e:
                                print(f"Error processing row {row}, col_idx {col_idx}: {e}")

                #Success message when results are complete and displays where the file was saved
                messagebox.showinfo("Success!", "Audit comparison results are complete and saved as 'Audit_Results.xlsx'.")
                return True
        
            except Exception as e:
                messagebox.showerror("Error", f"An error has occured: {str(e)}")
                return False
    
if __name__ =="__main__":
    root = tk.Tk()
    app = CompareFiles(root)
    root.mainloop()