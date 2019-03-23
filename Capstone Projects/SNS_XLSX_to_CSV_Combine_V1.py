#! python3
# 
# Converts all .xlsx files in the folder
# where the script is executed, to .csv, and 
# combines them into a single
# file. This script requires a command line 
# argument that represents
# the number of initial rows to be skipped
#
# Contact: a.sanchez.824@gmail.com
#
# V0: 3/5/2019
#		> Initial version ... 
#		Specific for SNS datasets
#
# V1: 3/20/2019
#		> Get rid of rows with "Total" in column 0
#		> Get rid of product description >> column 21
#		> reuse code from ARAP + SC ...
#
# 
#

#........................................................................
# libraries
#........................................................................
import xlrd 	# to work with xlsx files
import csv 		# to work with csv files
import os		# to work with directories, etc
import sys		# to get data from command line
import time		# for simple profiling
import re 		# to use regular expressions

#........................................................................
# functions
#........................................................................

# transforms numeric elements in 
# this_row (a list) that end in .0
# to integers
def integrize(this_row):
		integral_re = re.compile("^[0-9]+\.0{1}$")
		return [int(float(x)) if integral_re.match(str(x)) is not None else x for x in this_row]

# this was used to debug the previous function
# it turns that I was assuming that all elements
# in the_row were strings ... but the numeric ones
# are actually numeric!
def integrize_2(this_row):
		print("\n>> In Integrize_2: type(this_row) is: " + str(type(this_row)))
		integral_re = re.compile("^[0-9]+\.0{1}$")
		for x in this_row:
			print("\n>> In Integrize_2 (loop): type(x) is: " + str(type(x)))
			if integral_re.match(x) is not None:
				x = str(int(float(x)))
			else:
				continue
		return this_row

# for elements in this_row that are equals
# to garbage, transforms them to the empty
# string
def clean_garbage(this_row, garbage):
	return ["" if x == garbage else x for x in this_row]

# applies various cleaning methods to the
# the_row (a list)
def cleaned_row(this_row):
	#
	# gets rid of long dash
	# uses cute trick from ...
	# https://stackoverflow.com/questions/1540049/replace-values-in-list-using-python
	#GARBAGE = '—'
	#result = clean_garbage(this_row, GARBAGE)
	#
	# transforms numbers of the form xyz.0 to xyz using regex
	#result = integrize(result)
	result = integrize(this_row)
	#
	# end of cleanup
	#
	return result

# this function returns a string of the
# form yyyy-mm-dd, from the file_name
# that is assumed to have the form
# a_b_c where a, b, and c are strings
# and c has the form uuvvww, where
# uu represents the month, vv represents
# the day and ww represents the last digits
# of the year, such that the actual year
# is 2000 + int(ww)
def get_date(file_name):
	date_parts = file_name.split("_")
	the_date = date_parts[2]
	the_month = the_date[0:2]
	the_day = the_date[2:4]
	the_year = the_date[4:]
	result = str(2000 + int(the_year)) + "-" + the_month + "-" + the_day
	return result


# determines if given string is the value
# stored in the given column of the given
# list
def has_garbage(this_row, column_index, garbage):
	if this_row[column_index] == garbage:
		return True
	else:
		return False

# removes column given by its index 
def skip_column(this_row, column_to_skip):
	return this_row[0:column_to_skip] + this_row[column_to_skip+1:]

# given a file_name, whose extension is supposed
# to be xlsx, and which resides in directory
# dir_name, converts that file to csv in the
# same directory, with the same name ...
# also when converting the file, the first rows_to_skip
# rows are skipped ... finally, the cs_writer object
# is supposed to point to the output file that holds
# the concatenation of all files once transformed to csv
# code adapted from ...
# https://stackoverflow.com/questions/22688477/converting-xls-to-csv-in-python-3-using-xlrd
# NOTE: an extra empty line was being added after each line
# the solution of this problem is discussed here ...
# https://stackoverflow.com/questions/3191528/csv-in-python-adding-an-extra-carriage-return-on-windows
def xl_2_csv(file_name, dir_name, rows_to_skip, column_to_skip, csv_writer):
	FIRST_SHEET_INDEX = 0
	os.chdir(dir_name)
	print("\n>> [xl_2_csv] Processing directory: " + dir_name)
	xl_file_name = file_name + ".xlsx"
	print("\n>> [xl_2_csv] Processing file: " + xl_file_name)
	work_book = xlrd.open_workbook(xl_file_name)
	tha_sheet = work_book.sheet_by_index(FIRST_SHEET_INDEX)
	# get date in the form yyyy-mm-dd
	# from file_name
	#the_date = get_date(file_name)
	for row_number in range(rows_to_skip, tha_sheet.nrows):
		this_row = tha_sheet.row_values(row_number)
		# skip row if it contains 'Total' or '' 
		# in first column
		if (has_garbage(this_row, 0, 'Total') or 
			has_garbage(this_row, 0, '')):
			continue
		# filter out rows according to columns_to_keep
		this_row = skip_column(this_row, column_to_skip)
		# clean up row
		this_row = cleaned_row(this_row)
		csv_writer.writerow(this_row)
	print("\n>> [xl_2_csv] Converted file: " + file_name + ".csv")
	return

#........................................................................
# main logic
#........................................................................

# set initial clock tick
start_time = time.time()

# check that the following is provided
# 	> name of script
#	> number of top rows to be skipped
#	> list of columns to be kept in the output (at least 1)
num_arguments = len(sys.argv)
if (num_arguments < 3):
	print("\n***Error, please include the number of initial rows")
	print("to skip, and the index of single column to be skipped in the output")
	sys.exit("Script aborted!")

# get number of rows 
# to skip from command line
rows_to_skip = int(sys.argv[1])

# get the list of columns to be kept
# all the ones from indices 2 to num_arguments - 1
column_to_skip = int(sys.argv[2])

# setup output file
csv_file = open("combined_file.csv", 'w', encoding='utf8', newline='')
csv_writer = csv.writer(csv_file, quoting=csv.QUOTE_ALL)

# iterate over all files in current directory with extension .xlsx
current_directory = os.getcwd()
file_counter = 0
for xl_file_name in os.listdir('.'):
	if not xl_file_name.endswith('.xlsx'):
		continue
	xl_unqualified_name = os.path.splitext(xl_file_name)[0]
	xl_2_csv(xl_unqualified_name, 
		     current_directory, 
		     rows_to_skip, 
		     column_to_skip,
		     csv_writer)
	file_counter = file_counter + 1

print("\n>> [xl_2_csv] Number of converted files: " + str(file_counter))

# close output file
csv_file.close()

# report rough processing time
print("--- %s seconds ---" % (time.time() - start_time))

# TBTG!!!