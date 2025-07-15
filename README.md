# BOM-filter-Registor-and-Capacitor-
BOM filter Registor and Capacitor 

Resistor Filter Tool
This Python script automatically filters out components labeled as resistor from the PCBA worksheet in Geo_Diversity.xlsm and exports the relevant part numbers and descriptions for easy reference and further analysis.

Input File
Geo_Diversity.xlsm
Source Excel file containing the PCBA component list.

Output File
step1.xlsx
The filtered result including the following columns:

Type: Component type (e.g., resistor)

PN: Part Number

Description: Component description

How to Use
Make sure Geo_Diversity.xlsm is placed in the same directory as the script.

Run the Python script:


python your_script_name.py
The script will generate step1.xlsx in the same directory upon completion.

Logic Summary
Reads the PCBA worksheet from the input Excel file.

Column mapping:

Column C → LIST.2 (Type)

Column E → Unnamed: 4 (PN)

Column K → TEXT.4 (Description)

Filter condition:
Select rows where Type contains the word “resistor” (case-insensitive).

Step 2: Update Part Numbers Based on EE Change Mapping
This step takes the output from step1.xlsx and performs part number updates based on a change list defined in DCR0020430_checklist_0625.xlsx (EE Change sheet).

Input Files
step1.xlsx: Output from Step 1.

DCR0020430_checklist_0625.xlsx: Contains a mapping of old part numbers to new part numbers.

Column A: Changed list (Old PNs)

Column B: New PNs

Replacement Rule
If the 4th last character of a part number in the PN column is "S" and it exists in the mapping list, the PN will be replaced with the corresponding new PN from the checklist.

Output File
step2.xlsx: Contains original columns plus a new column:

Updated_PN: Updated part number (if applicable), or original if no match.

Logic Summary (Step 2)
For each part number in PN:

Check if the fourth-last character is S.

If so, and it matches an entry in the EE Change table, replace it.

Save updated data into step2.xlsx.

pandas
openpyxl (required to write .xlsx files)

Install them via pip:
pip install pandas openpyxl


