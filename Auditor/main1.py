
# standard imports
import re
from tkinter import filedialog, messagebox
import datetime
from tkinter import *
import os

# third parrty imports
import pandas as pd


# Global state to hold loaded dataframes so they can be accessed across functions
app_state = {
    "operator_df": None,
    "report_df": None
}

# Explicit mapping between report and operator columns allowing comparison for differing headers
column_mapping = {
    'Everi Game ID': 'Game',
    'Denom': 'Denom Selection',
    'Line Selection': 'Line/Ways Selection',
    'Bet Multiplier Selection': 'Bet Multiplier Selection',
    'Default Denom': 'Default Denom Selection',
    'Default Line': 'Default Line/Ways',
    'Default Bet Multiplier': 'Default Bet Multiplier',
    'Default Bet': 'Total Default Bet',
    'Min Bet': 'Min Bet',
    'Max Bet': 'Max Bet'
}
# get the operator column names directly from the mapping
operator_columns = list(column_mapping.values())

def normalize_game_id(game_id):
    # Normalzie game ids so they can all be compared fairly

    if pd.isna(game_id):
        return ""
    
    # Convert to string, lowercase, strip spaces and remove apostrophes and punctuation
    return (
        str(game_id)
        .strip()
        .lower()
        .replace(" ", "")         # Remove spaces
        .replace("'", "")         # Remove apostrophes
        .replace("’", "")         # Also handle curly apostrophes (from copy-paste)     
        .replace(":", "")         #And colons like for off the hook fishing     
    )



def normalize_value(value):
    if pd.isna(value):
        return []

    if isinstance(value, (int, float)):
        return [round(float(value), 2)]  # round for consistent comparison

    if isinstance(value, str):
        parts = re.split(r'[,\s]+', value)
        cleaned_numbers = []
        cleaned_strings = []

        for part in parts:
            part = part.strip()
            if not part:
                continue
            try:
                # Remove $ signs and convert to float
                num = float(part.replace('$', ''))
                cleaned_numbers.append(round(num, 2))  # rounding to avoid float precision issues
            except ValueError:
                cleaned_strings.append(part.lower())

        # Sort numeric and string parts separately, then combine
        return sorted(cleaned_numbers) + sorted(cleaned_strings)

    return [str(value).strip().lower()]


def load_operator_sheet():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not file:
        return

    # Step 1: Read first 10 rows without headers to scan for known column names - This can be increased later without performance hit
    preview_df = pd.read_excel(file, header=None, nrows=10)

    header_row = None
    # Loop through rows looking for a row that contains 'Game' (a key operator column)
    for i, row in preview_df.iterrows():
        if 'Game' in row.values:
            header_row = i
            break

    # If header row is not found, throw an error
    if header_row is None:
        messagebox.showerror("Error", "Could not find header row in operator sheet.")
        return

    # Step 2: Read actual data starting at detected header row
    df = pd.read_excel(file, skiprows=header_row)
    df.columns = df.columns.str.strip()  # Remove leading/trailing spaces in headers

    # Ensure 'Game' column exists before proceeding
    if 'Game' not in df.columns:
        messagebox.showerror("Error", "'Game' column not found in operator sheet.")
        return

    # Strip whitespace and ensure consistent string formatting for Game names
    df['Game'] = df['Game'].apply(normalize_game_id)

    # Remove duplicate Game rows to avoid false mismatches
    df = df.drop_duplicates(subset='Game')

    # Save cleaned DataFrame into app state
    app_state["operator_df"] = df
    messagebox.showinfo("Success", "Operator sheet loaded!")


def load_report_sheet():
    file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if not file:
        return

    # Step 1: Read the first 10 rows without assuming headers
    preview_df = pd.read_excel(file, header=None, nrows=10)

    header_row = None
    for i, row in preview_df.iterrows():
        if 'Game ID' in row.values or 'Denom' in row.values:
            header_row = i
            break

    if header_row is None:
        messagebox.showerror("Error", "Could not find header row in report sheet.")
        return

    # Step 2: Read the file starting from the correct header row
    df = pd.read_excel(file, skiprows=header_row)
    df.columns = df.columns.str.strip()

    # Step 3: Rename and clean up
    df = df.rename(columns=column_mapping)
    # print("[DEBUG] Columns after renaming:", df.columns.tolist())

    if 'Default Line/Ways' in df.columns:
        df['Default Line/Ways'] = df['Default Line/Ways'].astype(str).str.strip()

    if 'Game' not in df.columns:
        messagebox.showerror("Error", "'Game' column not found after renaming.")
        return

    df['Game'] = df['Game'].apply(normalize_game_id)    
    
    df = df.drop_duplicates(subset='Game')

    app_state["report_df"] = df
    messagebox.showinfo("Success", "Report sheet loaded!")


def run_audit():
    operator_df = app_state.get("operator_df")
    report_df = app_state.get("report_df")

    if operator_df is None or report_df is None:
        messagebox.showwarning("Missing Data", "Please load both sheets first!")
        return

    differences = []

    for _, operator_row in operator_df.iterrows():
        game_id = operator_row['Game']
        matching_row = report_df[report_df['Game'] == game_id]

        if matching_row.empty:
            differences.append({
                'Game': game_id,
                'Field': '',
                'Operator Value': '',
                'Report Value': '',
                'Issue': 'Missing in Report Sheet',
                'Source': 'Operator'
            })
            continue

        matching_row = matching_row.iloc[0]

        for col in operator_columns[1:]:
            op_val = normalize_value(operator_row[col])
            rep_val = normalize_value(matching_row[col])

            if op_val != rep_val:
                # print(f"DEBUG Mismatch for {col} | Operator: {op_val} | Report: {rep_val}")
                differences.append({
                    'Game': game_id,
                    'Field': col,
                    'Operator Value': operator_row[col],
                    'Report Value': matching_row[col],
                    'Issue': 'Mismatch',
                    'Source': 'Both'
                })

    # Also check for games in report but missing in operator
    operator_game_ids = set(operator_df['Game'])
    report_game_ids = set(report_df['Game'])

    missing_in_operator = report_game_ids - operator_game_ids

    for game_id in missing_in_operator:
        differences.append({
            'Game': game_id,
            'Field': '',
            'Operator Value': '',
            'Report Value': '',
            'Issue': 'Missing in Operator Sheet',
            'Source': 'Report'
        })

    if not differences:
        messagebox.showinfo("Audit Complete", "No differences found!")
        return

    # Convert to DataFrame and save
    diff_df = pd.DataFrame(differences)
    diff_df = diff_df[['Game', 'Field', 'Operator Value', 'Report Value', 'Issue', 'Source']]
    diff_df = diff_df.sort_values(by=['Game', 'Issue', 'Field'], ignore_index=True)

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             initialfile=f"Audit_Report_{timestamp}.xlsx",
                                             filetypes=[("Excel Files", "*.xlsx")])
    if save_path:
        diff_df.to_excel(save_path, index=False)
        messagebox.showinfo("Success", f"Audit complete! Report saved:\n{save_path}")
        import subprocess
        import platform

        # Open the file automatically after saving
        if platform.system() == 'Windows':
            os.startfile(save_path)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.call(['open', save_path])
        else:  # Linux or others
            subprocess.call(['xdg-open', save_path])


# ---------- UI handling ---------- #

window = Tk()
window.title("Auto Project")
window.config(padx=50, pady=50)

#pic_dir = os.path.dirname("AutoProject.png")
#img_path = "C:/Users/seth.jamieson/Desktop/Auditor/AutoProject.png"

# safe robust way to get pathing for whatever logo is chosen and to get around any weirdness.
logo_placeholder = "AutoProject.png"
pathtopic = os.path.join(os.path.dirname(__file__),logo_placeholder)
logo_photo = PhotoImage(file=pathtopic)
#logo_photo = PhotoImage(file="AutoProject.png") #old way to set logo photo but was somehwat broken

canvas = Canvas(width=500, height=500)
canvas.create_image(250, 250, image=logo_photo)
canvas.grid(column=1, row=0)

# BUTTONS #
upload_operator_button = Button(text="FIRST Operator sheet", width=24, command=load_operator_sheet)
upload_operator_button.grid(column=0, row=4) # sticky=E + W if needed Ugh, I dont remember how I found the sticky=E + W fix
upload_report_button = Button(text="Second Report sheet", width=24, command=load_report_sheet)
upload_report_button.grid(column=1, row=4) # sticky=E + W Ugh, I dont remember how I found the sticky=E + W fix
audit_it_button = Button(text="Third: Run Audit", width=24, command=run_audit)
audit_it_button.grid(column=2, row=4)

window.mainloop()
