# DigAutoProject
Test run of basic functionality for the auto audit project. 

The idea is for the program to intake a user selected excel sheet like an operator sheet and then intake another user selected excel sheet like a report file to compare against and audit. It should then produce a report excel file which highlights differences.

MVP:
Load both sheets,
Normalize Game IDs,
Normalize Data values,
Compare row by row per game ID,
Compare columns of interest,
output results to excel,
open excel file,
Simple UI
 
Nice to haves:
More robust error handling,
Better formatting in excel output(maybe color highlights etc),
Option to configure comparison fields - only compare or show default denoms or RTPs,
Possible integration with admin panel?


Created in python 3.10, pandas and tkinter

Known Collaborators:
Seth and Andrea
