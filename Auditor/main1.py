import sqlite3
from tkinter import *
import tkinter
from tkinter import filedialog
import pandas as pd

# Setting this up so that the user-selected file can be accessed globally
user_selected_file = None


### Originally was using sqlite3 and a series of queries to handle the excel files and audit logic.
### Removed until I am more familiar. Will try to re acquaint myself with pandas/dataframes
# # Database name
# connection = sqlite3.connect("audit_test.db")
# cursor = connection.cursor()

# When user presses button this should handle the initial file upload. SHOULD BE OPERATOR SHEET FORMAT
def user_first_upload():
    # Global variable to be used for user-selected file so we can access it easily in other functions
    global user_selected_file
    # Prevents an empty tkinter window from appearing
    tkinter.Tk().withdraw()
    # Produces a file selection window and sets settings
    user_selected_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])


# Merged the second file upload and audit function. When user presses second button it should prompt another file upload
# SHOULD BE ACTIVE DATA FORMATTED EXCEL
def audit_function():
    # Ask the user to upload the second Excel file
    second_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")])

    # Check if both files are selected
    if not user_selected_file or not second_file:
        print("Please select both Excel files.")
        return

    # Columns to compare in the first Excel file
    columns_to_compare1 = ['Game', 'Denom Selection', 'Line/Ways Selection', 'Bet Multiplier Selection',
                           'Default Denom Selection', 'Default Line/Ways', 'Default Bet Multiplier',
                           'Total Default Bet', 'Min Bet', 'Max Bet']

    # Columns to compare in the second Excel file (Might need to adjust. Also there is a much better way to do this...)
    columns_to_compare2 = ['Game', 'Denom', 'Line/Ways', 'Bet Multiplier', 'Default Denom', 'Default Line/Ways',
                            'Default Bet Multiplier', 'Total Default Bet', 'Min Bet', 'Max Bet']

    # Read the first Excel file into a DataFrame, skipping empty rows - retouch this later
    excel_df1 = pd.read_excel(user_selected_file, skiprows=lambda x: x in [0, 1])

    # # Troubleshooting to check which columns in columns_to_compare1 exist in the DataFrame
    # found_columns = [col for col in columns_to_compare1 if col in excel_df1.columns]
    # if not found_columns:
    #     print(f"No columns found in the first Excel file.")
    #     return

    # Read the second Excel file into a DataFrame
    excel_df2 = pd.read_excel(second_file)

    # Create a new DataFrame for the differences...
    differences = pd.DataFrame(columns=columns_to_compare1 + ['Differences'])

    # Iterate through games in the second Excel file
    for index, row in excel_df2.iterrows():
        game = row['Game']

        # Find the corresponding row in the first Excel file
        matching_row = excel_df1[excel_df1['Game'] == game]

        # Check if the game exists in both files (NEED TO EXPAN LOGIC FOR MISSING GAMES)
        if not matching_row.empty:
            # Create a dictionary to store differences
            diff_dict = {'Game': game}

            # Compare specified fields
            for col1, col2 in zip(columns_to_compare1, columns_to_compare2):
                # Check if the column exists in excel_df2
                if col2 in excel_df2.columns:
                    value1 = matching_row[col1].iloc[0]
                    value2 = row[col2]

                    # Check for differences
                    if value1 != value2:
                        diff_dict[col1 + '_diff'] = f"{value1} (File 1) != {value2} (File 2)"
                    else:
                        diff_dict[col1 + '_diff'] = ''

            # Append the differences to the DataFrame
            differences = pd.concat([differences, pd.DataFrame([diff_dict])], ignore_index=True)

    # Output the differences to a new Excel file (Could get fancy with color coding...or KISS)
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    differences.to_excel(output_file, index=False)

    print(f"Differences found and saved to {output_file}")


# ---------- UI handling ---------- #

window = Tk()
window.title("Auto Project")
window.config(padx=50, pady=50)

logo_photo = PhotoImage(file="AutoProject.png")
canvas = Canvas(width=500, height=500)
canvas.create_image(250, 250, image=logo_photo)
canvas.grid(column=1, row=0)

# BUTTONS #
upload_button = Button(text="FIRST Select operator excel sheet", width=36, command=user_first_upload)
upload_button.grid(column=1, row=4, columnspan=2, sticky=E + W) # Ugh, I dont remember how I found the sticky=E + W fix
search_button = Button(text="Select active data and audit", command=audit_function)
search_button.grid(column=2, row=1, sticky=E + W)

window.mainloop()