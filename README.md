# SAP-Excel-data-comparison
Macro written at work to make daily control for audit easier and faster.

Private project written to facilitate certain processes, sensitive data removed or changed.
A macro in an Excel file that is templated for daily duties when assigned to two buttons greatly speeds up work, minimizes the chance of error and makes it easier to implement for newcomers, by being intuitive and simple.

The two main functions which are copyAll() and createFolderAndSaveFiles() make it possible:
> Retrieve data and titles from the selected Excel file,
> Retrieve the relevant data for a given day from SAP,
> Compile in the corresponding spreadsheets the data using vlookup formulas,
> Identifying the differences if any, filtering them and adding a comment tab, if there are no differences, the program will disable the filters and return the information that there are none,
> After the user adds comments, the next button creates a folder with the appropriate suggested name in the corresponding folder,
> Selecting the relevant area setting it as print area and printing it to pdf to this folder (with appropriate properties: horizontal orientation, fit all columns on one page)
> Saving a copy of this Excel template in this folder.

Planned improvement:
> adding a function to compare differences with previous days, which will make it easier to fill in comments
