Hello and welcome!
I'd like to show you my "helping hand" - script made for a colleague to make his life easier. 

Script simply removes content from Excel file, based on another Excel file. At the moment it maintains only 'A' column (as it was necessary).

Script made in Python.

Packages used by script: openpyxl, os, sys.

You should name the database "db.xlsx" and the database to be deleted as "del.xlsx". Files should be located on PATH folder of script. Both databases should have records in column A in sheet "Sheet1". 
As a result you will receive Excel file named "db_del.xlsx" with only removed contents of database.

Enjoy!
