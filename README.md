excel2json
==========

Copyright (C) 2013 by Hojin Choi <hojin.choi@gmail.com>

Excel2json is a converting script that supports to managing well structured excel data to json format.

You can freely distribute this product with A-CUP-OF-BEER License (See source code)

*USAGE*

".js" extension files are associated with WSCRIPT.EXE, you can easily run the script by double click!

You may also make your own start script, like an 'excel2json.bat' with which you can run the script
specifying excel files and output folder name as the arguments.

	WSCRIPT.EXE Excel2Json.js file1.xlsx file2.xlsx product

*HOWTO-WORK*

By clicking the script in explorer:

	1. Run the script in a folder without any argument (by clicking)
	2. The script searches the folder for all excel files with extension .xls, .xlsx.
	3. All the sheets in the excel file are converted to CSV files.
	4. The CSV files are stored temporary folders with suffix (.$$$)
	6. Parse the CSV files and make json files into the 'output' folder.
	7. All the temporary folders will be removed with their contents (csv files)
	
By running WSCRIPT.EXE Excel2Json.js file1.xlsx file2.xlsx product:

	1. All the proceess is same with above.
	2. But it does not search the directory for excel files.
	3. And use the 'product' directory instead of 'output' for its result.
	
Excel-contents-format:

	See sample excel files! (Provided English, Korean versions)
	
