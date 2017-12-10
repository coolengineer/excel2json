excel2json
==========

Copyright (C) 2013 by Hojin Choi <hojin.choi@gmail.com>

Excel2json is a converting script that supports to manage well structured excel data to json format.

You can freely redistribute this product with A-CUP-OF-BEER License (See source code)

# USAGE #

## ADD SOME HINTS TO YOUR EXCEL FILE FOR JSON ##

There are four types of object, which can be handled by this script.

### Simple Object ###

If you have an excel file, very simple one like this.

|   |      A      |    B    |    C    |    D    |       E       |
|---|-------------|---------|---------|---------|---------------|
| 1 |             |  Initial  data    |         |               |
| 2 |             |   Name  |  Value  |         |               |
| 3 |             |  coins  |  1000   |         |               |
| 3 |             |  golds  |     0   |         |               |

And, Let's give a hint for the script, which awares A Column's '#' mark.

|   |      A      |    B    |    C    |    D    |       E       |
|---|-------------|---------|---------|---------|---------------|
| 1 |             |  Initial| data    |         |               |
| 2 |             |   Name  |  Value  |         |               |
| 3 | #initdata{} |   $key  |  $value |         | *inserted!*   |
| 4 |             |  coins  |  1000   |         |               |
| 5 |             |  golds  |     0   |         |               |

In the above example, you can get this JSON file

<pre>
{
	"initdata" : {
		"coins" : 1000,
		"golds" : 0
	}
}
</pre>

### Objects in Object ###
Above example explains plain value object, now if you want an object which
has objects as the key/value pairs, you can use "{{}}" suffix instead of "{}"

|   |        A       |      B     |    C    |    D    |       E       |
|---|----------------|------------|---------|---------|---------------|
| 1 |                |  Buildings |         |         |               |
| 2 |                |   Name     |  Color  |  Width  |    Height     |
| 3 | #buildings{{}} |   $key     |  color  |  width  |    height     |
| 4 |                |  barrack   |  blue   |   200   |     200       |
| 5 |                |  mine      |  yellow |   200   |     100       |
| 6 |                |  gas       |   red   |   100   |     100       |
| 7 |                |  townhall  |  black  |   200   |     200       |

And, this yields
<pre>
{
	"buildings" : {
		"barrack" : {
			"color": "blue",
			"width": 200,
			"height": 200
		},
		"mine" : {
			"color": "yellow",
			"width": 200,
			"height": 100
		},
		"gas": {
			"color": "red",
			"width": 100,
			"height": 100
		},
		"townhall": {
			"color": "black",
			"width": 200,
			"height": 100
		}
	}
}			
</pre>

### Arrays in Object ###

This type of object has nested value as an array, see this example!

|   |        A       |      B     |    C    |    D    |   E   | 
|---|----------------|------------|---------|---------|-------|
| 1 |                |  Required coins of buildings   |       |
| 2 | #reqcoins{[]}  |   barrack  |  mine   |  gas    |       |
| 3 |                |        100 |    100  |   100   |       |
| 4 |                |        500 |    500  |   500   |       |
| 5 |                |       1000 |   1000  |  1000   |       |
| 6 |                |       1500 |         |         |       |

As you can see, the suffix of #reqcoins is "{[]}", this gives hints for constructing vertical array.
The result is

<pre>
{
	"reqcoins" : {
		"barrack" : [100, 500, 1000, 1500 ],
		"mine" : [ 100, 500, 1000 ],
		"gas"  : [ 100, 500, 1000 ]
	}
}
</pre>

### Object Array ###

The last format of compound data is an array which contains objects, the suffix "[{}]"

|   |     A     |      B     |    C    |    D     |   E   | 
|---|-----------|------------|---------|----------|-------|
| 1 |           |    Shop    |         |          |       |
| 2 | #shop[{}] |   name     |  price  | category |       |
| 3 |           |    blade   |    100  |  attack  |       |
| 4 |           |    dagger  |    200  |  attack  |       |
| 5 |           |    shield  |    100  |  defese  |       |

And this yields

<pre>
{
	"shop" : [
		{
			"name": "blade",
			"price" : 100,
			"category" : "attack"
		},
		{
			"name": "dagger",
			"price": 200,
			"category": "attack"
		},
		{
			"name": "shield",
			"price": 100,
			"category": "defense"
		}
	]
}
</pre>

### Array Value (Tip) ###

Magic field suffix "[]" of object description line introduce array value.

|   |        A       |      B     |             C         |    D     |   E   | 
|---|----------------|------------|-----------------------|----------|-------|
| 1 |                |    Shop    |                       |          |       |
| 2 | #inventory[{}] |    type    |        attrib[]       |    dur   |       |
| 3 |                |    blade   |       oil, diamond    |    100   |       |
| 4 |                |    dagger  |         sapphire      |    150   |       |
| 5 |                |    shield  | diamond,sapphire,rune |    200   |       |

The "attrib[]" field name terminates with "[]", which indicates attrib key has
array value. so, the result will be like this.

<pre>
{
	"inventory" : [
		{
			"type": "blade",
			"attrib": [ "oil", "diamond" ],
			"dur": 100
		},
		{
			"type": "dagger",
			"attrib": [ "sapphire" ],
			"dur": 150
		},
		{
			"type": "shield",
			"attrib": [ "diamond", "sapphire", "rune" ],
			"dur": 200
		}
	]
}
</pre>

## Some technical marks ##

### '!' prefixed sheet name ###

You can insert '!' mark before a sheet name which will not be considered to be parsed. For e.g. '!Samples', '!Test' or '!Templates'.

## RUN Excel2Json.js ##

".js" extension files are associated with WSCRIPT.EXE (Windows already has this program!),
you can easily run the script by double click!

You may also make your own start script, like an 'excel2json.bat' with which you can run the script
specifying excel files and output folder name as the arguments.

	MKDIR output
	WSCRIPT.EXE Excel2Json.js file1.xlsx file2.xlsx product

# HOWTO-WORK #

Internally, CSV format is used to parse excel files.

By clicking the script in explorer:

	0. Make 'output' folder (mkdir output)
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
	
