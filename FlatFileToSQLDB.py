import csv
import pyodbc
# Insert the file path for the file you are looking to put into SQL DB
File = 'InsertFilePathHere'
#Insert the name of the table within your SQL DB you want to put data in (must have table created for this code)
Table = 'InsertTableNameHere'

# Creating variables for the connection string later in the script
con = pyodbc.connect('Driver={SQL Server};Server=SERVERNAMEHERE;Database=DATABASENAMEHERE;Trusted_Connection=yes;')
cursor = con.cursor()


def queryExecution(file, table):
		infile = open(file, 'r')
		lines = infile.readlines()

		# Depending on your file you may need to strip other characters or split on a different delimiter
		# Also a good place to clean up the text in the file adding additional .strip or .replace
		colArray = lines[0].strip('\n').split(',')

		# Building colum string for SQL statement
		colStr = ''
		for column in colArray:
			colStr = colStr + "[" + column + "]" + ","
		colStr = '(' + colStr.rstrip(',') + ")"

		# Building a DictArray structured [Column1:Row1].... and so on
		# Also a good place to clean up the data within the rows
		rowArray = []
		for row in lines[1:]:
			row = row.split(',')
			dict = {}
			rowIndex = 0
			for i in row:
				dict[colArray[rowIndex]] = i
				rowIndex += 1
			rowArray.append(dict)

		# The building of the SQL Statement
		# The truncate table statement was added because we run this script every night and repopulate for some tables
		truncate = 'TRUNCATE ' + 'TABLE ' + table
		# Executing the truncate statement to set us up with a nice and empty table (if not already)
		cursor.execute(truncate)
		con.commit()

		# Building the values portion of our SQL Statement
		# We can not insert None type so we transfor to NULL in for loop
		# If you need help visualizing whats happening it helps to use print statements where appropriate
		valStr = " VALUES "
		count = 0
		for row in rowArray:
			rowStr = '('
			for column in colArray:
				if row.get(column) != None and row.get(column) != '':
					rowStr = rowStr + "'" + row.get(column).replace("'", '') + "'" + ","
				else:
					rowStr = rowStr + "NULL,"

			# Cleaning up the insert statement
			rowStr = rowStr.rstrip(',')
			valStr = valStr + rowStr + ')' + ","

			# We inserted the count variables because SQL can only handle inserting 1000 rows from an insert statement at one time
			# So every 900 rows we execute the statement
			count += 1
			if count == 900:
				valStr = valStr.rstrip(',')
				query = "INSERT INTO " + table
				query = query + colStr + valStr
				cursor.execute(query) #execute query
				con.commit() #commit changes
				valStr = " VALUES "
				count = 0
		# This is to catch the remaining values when there is less than 900 remaining
		if count > 0: #if vales still exist in dictarray loop back through for another print statement
			valStr = valStr.rstrip(',')
			query = "INSERT INTO " + table
			query = query + colStr + valStr
			cursor.execute(query) #execute query
			con.commit() #commit changes
			valStr = " VALUES "
			count = 0




print('Start')
queryExecution(File, Table)
print('End')
