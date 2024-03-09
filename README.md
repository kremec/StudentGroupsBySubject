# Student Groups By Subject

This program is designed to take a list of students and grouping them so each team has one member from each subject group, while supporting having exceptions for students who cannot work together.

## Usage

Start the program by running the python file.\
The program will display a file dialog to select the Excel file to use.

The file should be formatted as follows:
- First sheet: names of students in columns based on their subjects
- Second sheet: names of students in columns based on who they cannot work with

Names of sheets are not important.\
On both sheets, the first row is not included in the data to be parsed, as it is assumed to be a header row.

The program will then display the second file dialog to save the Excel file with generated student groups, each row representing one team.\
The file will be then opened automatically.