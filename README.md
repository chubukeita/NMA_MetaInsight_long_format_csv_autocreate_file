# NMA_MetaInsight_long_format_csv_autocreate_file
One software to perform Network Meta-analysis is MetaInsight (https://crsu.shinyapps.io/MetaInsight/). This requires uploading a csv file in either Long format or Wide format format. However, it is labor-intensive to manually input data and create csv files in these formats for the number of outcomes of interest. Therefore, we created a format that is more intuitive and easier to input data, and implemented a macro effective book that converts the data into Long format csv files based on the input data, after refactoring with ChatGPT 4.0 (hereinafter referred to as the Input Tool). The Input Tool allows the creation of Long format csv files at the touch of a button.

## initialization
Please unzip the Zip file (this zip file contains the macro-enabled book (e.g. NMA_MetaInsight_long_format_csv_autocreate_file.xlsm)) and activate the macro-enabled book before use. And please enable "Selenium".

## User Form Layout
The layout of labels, text boxes, combo boxes, and command buttons in the user form, as well as control names, can be checked in the slide file titled "UserFormLayout.pptx."

## precautions
The input tool basically manipulates the "InputSheet" to input data. Yellow cells are basically cells that do not contain Excel formulas, so you can change the input values depending on the situation. Cells that are not colored are basically cells that contain Excel formulas, so do not edit the cells. Also, avoid inserting rows and columns for all sheets as this may corrupt the data.


## Basic usage (for use when there are no more than 8 groups to be compared in the NMA)
## *For cases where the total number of groups exceeds 8, please also refer to "How to use NMA when the total number of groups to be compared exceeds 8" in the README file.
To begin, the three command buttons placed on the "SortSheet" allow you to specify the type of group classification.

### Group (arm) classification with "SortSheet”
The label names for each group outputted in MetaInsight will correspond to the values entered in cells C3 to C10 of the "SortSheet". When clicking the "classification into single or combined groups" button, you can change the input values from cells D2, E2, and F2. When clicking the "classified into the usual 8 groups" or "classification by Threshold" buttons, you can directly change the input values from cells C3 to C10.

Clicking the "classification into single or combined groups" button allows you to distribute interventions into single or combined groups based on the values entered in cells D2, E2, and F2. In this state, the values in cells D2, E2, and F2 can be changed, thereby altering the arm labels in column C. The corresponding IndexA and IndexB values for each arm are indirectly specified by entering them in the "InputSheet".

Clicking the "classified into the usual 8 groups" button allows classification into arms from cells C3 to C10. Corresponding IndexA and IndexB values for each arm are indirectly specified by entering them in the "InputSheet".

Clicking the "classification by Threshold" button launches the "Threshold settings" form for classifying IndexA into 4 levels and IndexB into 2 levels. You can set thresholds sequentially in the "IndexA" and "IndexB" tabs. In the "IndexA" sheet, IndexA is classified into 4 levels according to the threshold values in the table (range from cell B4 to F9).

First, enter the value in cell E7 and select "less than" or "or below" from the list. This sets the upper threshold for "Low". Similarly, enter the value in cell E8 and choose "less than" or "or below". This sets the upper threshold for "Intermediate". Pressing the "IndexInput" button reflects the form information in the table on the "IndexA" sheet. Similarly, in the "IndexB" tab, enter the value in cell E8, choose "less than" or "or below", and press the "IndexInput" button to reflect the form information in the table on the "InputB" sheet. The values of the two parameters reported in each paper are considered as "IndexA" and "IndexB", and entering them in the "IndexA" and "IndexB" columns of the "InputSheet" sheet automatically classifies the arms according to the thresholds set in the "Threshold settings" form.

### Input in InputSheet
In the 3rd row of the "InputSheet", yellow cells can have their names changed, and outcome names can be altered from column Z onwards. Double-clicking on a table in the "InputSheet" displays the Input Form, which can be used for data entry. The currently editing row is highlighted in light blue. Generally, cells that can be edited are colored, while those without color, containing formulas, should not be edited. Tabs can be switched with "Ctrl+Tab" to move right and "Ctrl+Shift+Tab" to move left.

#### Information Tab
The leftmost tab in the Input Form (default titled "Information") allows for entering basic information about the study. "StudyNo" indicates the number of the study as it appears from the top in the "InputSheet." Additionally, "Authors," "PMID," "Year," "Country," and "Research Period" allow entry of the lead author of the paper, PubMed ID, publication year, country, and study duration, respectively. Essential input items for NMA in MetaInsight are "StudyNo" and "Authors”.

Clicking the "Back" and "Next" buttons allows you to move the selected cell one up or down, respectively. Pressing the "Update" button updates the contents entered in the form. It's important to note that if you move the selected cell without pressing "Update" after re-entering data in the form, the sheet will not be updated with the new values. Therefore, ensure to press the "Update" button after changing input values in the form. Additionally, pressing the "Outcome_Add" button displays the "Outcome Setting" form, allowing you to add any desired outcome to the far right end of the "InputSheet." If the targeted outcome is binary data, select "Dichotomous" from the list, and for continuous outcomes, select "Continuous".  You can add an outcome by entering the name of the outcome in the textbox labeled "outcome_name" and clicking the "Outcome_Insort" button. The "Add" button allows you to move to the next row below the last row where text has been entered, enabling the input of a new paper. If the table is fully filled to the last row, a new input field is automatically created one row below. The "RowDelete" button can delete the selected row within the table in the "InputSheet".

#### "Strategies" Tab
The second tab from the left in the Input Form (default titled "Strategies") allows for entering PEEP, VT, and sample sizes for each group in the study. It supports the input of studies comparing up to three groups, namely "treatment1," "treatment2," and "treatment3." The input form for treatment1 is vertically aligned with "IndexA1," "IndexB1," and "Patients (n1)" representing IndexA, IndexB, and sample size of the first group, respectively. The subscript numbers indicate to which group they belong. If any values are unknown, leave them blank. After entering "IndexA" and "IndexB",  clicking the "Update" button in the Information tab updates the treatment and also the names of each group (arm).

#### "Outcome" Tab
Tabs from the third leftward in the Input Form vary depending on whether the outcome is continuous or dichotomous. A common entry for each outcome is the sample size (n). The difference from "Patients (n)" entered in the "Strategies" tab is whether protocol deviators occurring during clinical trials are considered. If dropouts during the trial are considered, the sample size originally assigned to each group at the protocol stage may differ. Generally, "Patients (n)" entered in the "Strategies" tab matches, but if it differs, a different number should be entered in "n". You can change the number in "n" by checking "Change n of treatment." Checking this and pressing "Update" turns the cell of "Patients (n)" in the "input" sheet red for the selected outcome, and the original formula in the cell is overwritten with a value paste. If you accidentally check and press "Update", you can revert by copying a cell with a formula and no color fill from the same column and pasting it over the red cell.

##### Continuous Outcome
In the outcome tab for continuous outcomes, you can enter outcomes with mean (μ) and standard deviation (±SD). If any values are unknown, enter "NR" and do not leave them blank.

##### Dichotomous Outcome
In the outcome tab for dichotomous outcomes, you can enter the number of events. If any values are unknown, enter "NR" and do not leave them blank.

#### "Create Sheets for Each Outcome" Button
Once all data is entered, press the "create sheets for each outcome" button at the bottom of the InputSheet table to convert the data into the Long format used in MetaInsight. When you press the button, a warning saying, "Before executing this macro, please save all necessary data. When the macro is executed, the original data will be lost. Have you saved your data?" will appear. Click "OK" to proceed. Once the message "complete" appears and you press "OK," sheets titled "(outcome name) table" will appear to the left of the input sheet. These sheets are a collection formatted in MetaInsight's Long Format. In each table sheet, comments are attached from cell A2 onwards in column A, lining up "StudyNo", "Authors", "PMID" and "Year" from the input sheet for verification.

#### "CSV File Output" Button
The "csv file output" button can convert all sheets to the left of the InputSheet into csv files. When you press the button, a warning stating, "Executing this macro will result in the loss of data in the xlsm file. Please save the xlsm file and copy the file before creating the csv file. The csv file output should be done after copying the xlsm file. Have you copied it?" will appear. Follow the warning to save the xlsm file once and duplicate it by copying the file. Then, proceed with the same operation on the duplicated file and click "OK" in response to the warning. When the guidance "Scope to csv" appears, enter "1" for start and the number of outcomes for end to output all outcomes as csv. This will open a dialog box titled "Specify the file to save" where you can save the file in any location. By default, the csv file is saved as "(outcome name) table.csv," and if a csv file with the same file name already exists in the saved location, the new file will be saved as "(outcome name) table (2).csv", "(outcome name) table (3).csv" and so on.

#### Other Buttons
Pressing the "Link_List create" button automatically creates URLs in the "Link_List" sheet for accessing the PubMed site of each paper. The "Connecting PMIDs with ORs" button can create a search formula linking all PMIDs entered in the input sheet with ORs. This search formula is displayed in cell A2 of the "Link_List" sheet. However, be cautious as it automatically operates Chrome, so do not manually operate the browser during the automation. This OR-linked search formula is necessary to check whether a self-made search formula for SR includes all papers in the same CQ as the prior SR. Specifically, checking if the total number of search results from the AND-linked search formula of "the OR-linked search formula of PMIDs collected in the prior SR" and "the self-made search formula" matches the number of papers collected in the SR confirms whether a comprehensive search was achieved. If the number of papers does not match, remake the search formula, repeat the process, and keep experimenting with the search formula until the numbers match.

The "delete the table on the left side of InputSheet" button can collectively delete the collection of sheets titled "(outcome name) table" created to the left of the InputSheet. The "hide uncolored columns" and "undo hiding of columns" buttons can hide and unhide cells without color on the InputSheet, respectively.

## How to use NMA when the total number of groups to be compared exceeds 8
If the total number of groups to be compared in NMA exceeds 8, please directly enter the names of the groups to be classified in rows 6 and onwards of columns W to Y (the "arm1" to "arm3" columns in the "Strategies" section) in the InputSheet. Delete and overwrite the formulas originally in rows 6 and onwards of columns W to Y with the names of the groups. However, entering data in the "IndexA" and "IndexB" of the "Strategies" tab in the "Input Form" that appears when double-clicking a cell in the InputSheet will no longer change the group names, but data entered for "Patient (n)" will be reflected in the InputSheet.

## License for this file
For this program (except Module4), this code was created by chubukeita and refactored by ChatGPT 4.0.

Copyright (c) chubukeita, subject to MIT License.

More information about the new license can be found at the following link: 

https://github.com/chubukeita/NMA_MetaInsight_long_format_csv_autocreate_file/blob/main/LICENSE

For Module4, the code for this module is taken from https://github.com/yamato1413/WebDriverManager-for-VBA.

Copyright (c) yamato1413, subject to the MIT License.

The full license can be found at the following link: https://github.com/yamato1413/WebDriverManager-for-VBA/blob/main/LICENSE

Changed from: https://github.com/yamato1413/WebDriverManager-for-VBA/blob/main/WebDriverManager4SeleniumBasic.bas and README '// SeleniumBasic ' in  and README.
