excel2json
==========

Copyright (C) 2013 by Hojin Choi <hojin.choi@gmail.com>

Excel2json is a converting script that supports to managing well structured excel data to json format.

You can freely distribute this product with MIT License

USAGE:
Run WSCIPT.EXE with an argument 'Excel2Json.js', do not make any html file for which include this js file
or, you can just double click, for .js extension files are associated with WSCRIPT.EXE, you can easily
run the script.

You may also make your own start script, like a 'excel2json.bat' with which you can run the script
specifying excel files and output folder name as the arguments. see below.

HOWTO-WORK:
By clicking the script in explorer:
	1. Run the script in a folder without any argument
	2. The script searches the folder for all excel files with extension .xls, .xlsx.
	3. All the sheets in the excel file are converted to CSV files.
	4. The CSV files are stored temporary folders named the excel file with additional suffix (.$$$)
	   (As many temporary folders as excel files will be created temporarily.)
	6. Json files are created in the 'output' folder if not existed then one will be created.
	7. All the temporary folders will be removed with their contents (csv files)
	
By running wscript.exe Excel2Json.js file1.xlsx file2.xlsx product (for e.g.)
    1. All the proceess is same with above.
	2. But it does not search the directory for excel files.
	3. And use the 'product' directory for its output storage.
	
Excel-contents-format:

	See sample excel files! (Provided English, Korean versions)