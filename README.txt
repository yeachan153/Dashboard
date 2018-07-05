DESCRIPTION:

This code was written as a group of 3. It was originally used to clean, merge excel files, data wrangle and give some descriptive statistics once an excel file with multiple 
sheets was given to it. The R File would then output a new csv file that would then be copied onto a google sheets dashboard. Most of the calculations were done in sheets to
make the dashboard interactive, thus the R File itself does not output a lot. Its main purpose is to wrangle data and make it future-proof for the client. I have linked the dashboard, however personal data has been made anonymous.

ORIGINAL README:

Steps that are required for the Dashboard:

1. Open the Dashboard Stats Store.r-file. 
2. Set your working directory using the setwd-function.
3. Run the entire R-script.
4. Open the created Dataset_Dashboard.csv-file in google sheets. 
5. Enter the Dashboard by copying the following link into your browser:

https://docs.google.com/spreadsheets/d/1L9d6HuRHHZF0zsjro73ojhiqSg2Rf-mc5k0tP6Av7YI/edit?usp=sharing

6. Copy the entire sheet from the Dataset_Dashboard.csv-file into the 'data'-sheet of the Dashboard-file. While copying, make sure your cursor is on A1.
7. Click the "Data" column on the top of the page, and select "Named ranges...". A right hand column will appear in your sheet. There will be two named ranges present: "Data" and "Filtered". Select "Data", click the pencil icon to edit. Leave the name of the named range as "Data", but ensure that your selected range spans the entire sheet. You may achieve this using the shortcut: "CTRL-A".

N.B. Further filters can be added by accessing the hidden sheets through View -> Hidden Sheets.
