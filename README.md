MS-Powerpoint-to-Excel-Table-Extractor

This program written in Python is designed to extract large quantities of complex data from tables (detected as shapes) from Microsoft Powerpoint making use of the pptx module. This data is then exported to MS Excel, in turn making use of the openpyxl module. 

The program, "power_to_excel.py" can export data from simple MS Powerpoint presentations where tables are assumed to be the same width i.e contain the same number of column headers. In turn, "power_to_excelv3.py" is able to extract data from multiple table-types. Notably, the number of different table types is detected first based on the uniqueness of a list of each table's column headers. Each list is termed a "header_template" and a dictionary is created to store these templates for comparison and reference. In addition, a separate sheet in MS Excel is created for each table-type or template. The presentation is then parsed and each table compared against the stored dictionary of templates. If the table is found to correspond to one of these templates, data is copied across to the corresponding sheet in the open MS Excel file. 

The program, "power_to_excelv4.5.py" performs the same task as "power_to_excelv3.py" except parses the Powerpoint presentation only once as oppose to twice i.e. it is more efficient code. 

The program, "power_to_excelv5.py" makes use of command-line arguments to receive input from the user i.e. which MS Powerpoint file they would like to make use of for transferring data to MS Excel. The program is also modularised into a class. 

The program, "power_to_excelv6.py" contains a class named "Combine" which inherits from the class "Powerpoint_to_Excel". This class overrides the main method so as to save the output Excel file as a different name and to run an additional method called "combine" which combines all data from all table templates into a single Excel sheet. In this version, data is combined by column, but unfortunately this does not take into account gaps and missing cells in tables and hence can lead to data becoming "bunched" in the combined version. 

The program, "power_to_excelv7.py" corrects this issue by combining data by row - zipping up each into a dictionary with its corresponding header template as keys. Matching keys are then searched for in the combined sheet header template (the latter selected for the combined sheet based on length) and when found the data copied across accordingly. 
