# This home-made modules contains functions to import data from various sources.
# In order to get Python to see the module, added following to .bash_profile or .bashrc:
#
# PYTHONPATH=$PYTHONPATH:/Users/phil/Dropbox/phjPythonModules; export PYTHONPATH
#
# To use the functions in this module, import phjReadData and refer to functions using:
# data = phjReadData.function_name(arg)
#
# e.g.
#	data = phjReadData.phjReadDataFromExcelNamedCellRange('/Users/phil/Dropbox/Clarkson/phjDataweatheranalysis.xlsx', '2011 data', 'tides_data')


def phjReadDataFromExcelNamedCellRange(phj_file_name, phj_sheet_name, phj_range_name, phj_datetime_format):
	# This function reads data from a named cell range in an Excel spreadsheet (.xlsx).
	# Input file_name (including path to file), sheet_name and range_name.
	# If the function can find the correct cell range, it reads the data and returns
	# it in the form of a pandas DataFrame.
	
	# Jan 2014
	# --------
	
	# IMPORT REQUIRED PACKAGES
	# ========================
	# The following method to check that a package is available is not considered to be
	# good practice:
	#
	#	try:
	# 		import package_name
	# 		HAVE_PACKAGE = True
	# 	except ImportError:
	# 		HAVE_PACKAGE = False
	#
	# Instead use the following method (as described at: http://developer.plone.org/reference_manuals/external/plone.api/contribute/conventions.html#about-imports)...
	
	import pkg_resources
	
	try:
		pkg_resources.get_distribution('openpyxl')
	except pkg_resources.DistributionNotFound:
		print "Error: openpyxl package not available."
		exit(0)
	else:
		import openpyxl
		from openpyxl.shared.exc import InvalidFileException


	try:
		pkg_resources.get_distribution('numpy')
	except pkg_resources.DistributionNotFound:
		print "Error: numpy package not available."
		exit(0)
	else:
		import numpy as np


	try:
		pkg_resources.get_distribution('pandas')
	except pkg_resources.DistributionNotFound:
		print "Error: pandas package not available."
		exit(0)
	else:
		import pandas as pd


	from decimal import Decimal

	
	# LOAD DATA FROM EXCEL SPREADSHEET
	# ================================
	try:
		# Load workbook. If unsuccessful, throws an InvalidFileException (imported in header)
		workbook = openpyxl.load_workbook(filename = phj_file_name, data_only=True, use_iterators = True)
		
	except InvalidFileException:
		print "File named '" + filename + "' does not exist."
		exit(0)
		
	# Look up worksheet by name. Function returns None if sheet of that name not found
	worksheet = workbook.get_sheet_by_name(phj_sheet_name)
		
	if worksheet is None:
		print "No such sheet in workbook."
		exit(0)

	# Get named range of cells. Function returns None if named_range not found...
	cell_range = workbook.get_named_range(phj_range_name)
	
	if cell_range is None:
		print "No such named range of cells."
		exit(0)
		
	print dir(cell_range)					# Get list of all iterable properties of object
	print cell_range.destinations			# This gives a list containing tuples of worksheet names and cell ranges 
	print cell_range.destinations[0]		# Give first tuple
	print cell_range.destinations[0][1]		# Gives element [1] of tuple [0] i.e. the cell range

	# Define list to store data...
	imported_data = []
	
	# Step through each row in cell range.
	# (N.B. Cells returned by iter_rows() are not regular openpyxl.cell.Cell but openpyxl.reader.iter_worksheet.RawCell.)
	for row in worksheet.iter_rows(cell_range.destinations[0][1]):

		# Define temporary list to store values from cells in rows
		tempdata=[]

		# Step through each cell in row...
		for cell in row:
			# print cell.internal_value		# Can use for debugging purposes.
			if cell.is_date:
				# If the cell contains a date, the format of cell.internal_value is, for
				# example, datetime.datetime(2011, 1, 1, 0, 0). Therefore, reformat to
				# using required format.
				# tempdata.append(cell.internal_value.strftime("%d/%m/%Y %H:%M:%S"))
				tempdata.append(cell.internal_value.strftime(phj_datetime_format))
	
			elif cell.data_type == 's':		# TYPE_STRING = 's' AND TYPE_STRING_NULL = 's'
				tempdata.append(cell.internal_value)

			elif cell.data_type == 'f':		# TYPE_FORMULA = 'f'
				# Including 'data_only=True' in openpyxl.load_workbook() means that formulae aren't recognised as formulae, only the resulting value.
				tempdata.append(cell.internal_value)

			elif cell.data_type == 'n':		# TYPE_NUMERIC = 'n'
				# Check if cell is empty...
				if cell.internal_value == None:
					# tempdata.append('Empty')
					tempdata.append(np.nan)
					
				else:
					# Some decimal values cannot be stored exactly. The decimal module
					# contains Decimal which helps with handling decimal values.
					# tempdata.append(Decimal(cell.internal_value).quantize(Decimal('1.00')))
					tempdata.append(cell.internal_value)
					
			elif cell.data_type == 'b':		# TYPE_BOOL = 'b'
				tempdata.append(cell.internal_value)

			elif cell.data_type == 'inlineStr':		# TYPE_INLINE = 'inlineStr'
				tempdata.append(cell.internal_value)

			elif cell.data_type == 'e':		# TYPE_ERROR = 'e'
				tempdata.append(cell.internal_value)

			elif cell.data_type == 'str':		# TYPE_FORMULA_CACHE_STRING = 'str'
				tempdata.append(cell.internal_value)

			else:
				tempdata.append("Unknown")
				
		imported_data.append(tuple(tempdata))
	
	variable_names = phjDealWithHeaderRow(imported_data)
	
	# Convert dataset to pandas DataFrame.
	# Each column now headed with original column headers as seen in Excel file
	# if header row present or with generic labels of 'var1', 'var2', etc. if no
	# header row present.
	phj_data_frame = pd.DataFrame(imported_data, columns=variable_names)
	
	print phj_data_frame

	return (phj_data_frame, variable_names)


def phjDealWithHeaderRow(data):
	# This function gets the variable names from the first row of data and
	# removes the header row from the data list. The data is passed by
	# references and, therefore, it can be mutated without having to make
	# a copy of the whole dataset.
	
	# Jan 2014
	# --------
	
	# Deal with headers in first row...
	header = raw_input("Does first row of data include header names? (y/n): ")

	if header == "y":
		# Identify variable names from first row of data...
		variable_names = data[0]
		
		# Remove header names from data...
		del data[0]
		
		print "First row (containing variable names) has been removed from the data."
		
	else:
		variable_names = []
		for i in range (len(data[0])):
			variable_names.append('var'+str(i+1))
			
	# Print variable names - just for reference...
	for i in range (len(variable_names)):
		print "var",i+1,": ",variable_names[i]
	
	print ''
		
	return variable_names



def phjDataReadFromStata(stataFileName):

#	4 Feb 2014
#	----------

# The following method to check that a package is available is not considered to be
# good practice:
#
#	try:
# 		import package_name
# 		HAVE_PACKAGE = True
# 	except ImportError:
# 		HAVE_PACKAGE = False
#
# Instead use the following method (as described at: http://developer.plone.org/reference_manuals/external/plone.api/contribute/conventions.html#about-imports)...
	
	import pkg_resources
	
	try:
		pkg_resources.get_distribution('numpy')
	except pkg_resources.DistributionNotFound:
		print "Numpy package not available."
		exit(0)
	else:
		import numpy as np
		
	try:
		pkg_resources.get_distribution('pandas')
	except pkg_resources.DistributionNotFound:
		print "Pandas package not available."
		exit(0)
	else:
		import pandas as pd

#	The Pandas I/O api is a set of top level reader functions accessed
#	like pd.read_csv() that generally return a pandas object.
#	(See: http://pandas.pydata.org/pandas-docs/stable/io.html)
#	
#	pandas.io.stata.read_stata
#	==========================
#	Taken from: http://pandas.pydata.org/pandas-docs/dev/generated/pandas.io.stata.read_stata.html
#	(Accessed 3 Feb 2014)
#	
#	pandas.io.stata.read_stata(filepath_or_buffer, convert_dates=True, convert_categoricals=True, encoding=None, index=None)
#	Read Stata file into DataFrame
#	
#	Parameters :	
#		filepath_or_buffer : string or file-like object
#			Path to .dta file or object implementing a binary read() functions
#		convert_dates : boolean, defaults to True
#			Convert date variables to DataFrame time values
#		convert_categoricals : boolean, defaults to True
#			Read value labels and convert columns to Categorical/Factor variables
#		ncoding : string, None or encoding
#			Encoding used to parse the files. Note that Stata does not support unicode. None defaults to cp1252.
#		index : identifier of index column
#			identifier of column that should be used as index of the DataFrame
	
	print stataFileName
	
	try:
		phjStataData = pd.read_stata(stataFileName, convert_dates=True, convert_categoricals=True, encoding=None, index=None)
	except:		# Catch all exceptions
#		e = sys.exc_info()[0]
		print "An error occurred reading data from Stata file."
		exit(0)

	print phjStataData
	return phjStataData



#################################
# Read data from an R workspace #
#################################
def phjReadDataFromR(phjPathToNewWorkingDirectory, phjRData):

#	4 Feb 2014
#	----------

# The following method to check that a package is available is not considered to be
# good practice:
#
#	try:
# 		import package_name
# 		HAVE_PACKAGE = True
# 	except ImportError:
# 		HAVE_PACKAGE = False
#
# Instead use the following method (as described at: http://developer.plone.org/reference_manuals/external/plone.api/contribute/conventions.html#about-imports)...
	
	import pkg_resources
	
	try:
		pkg_resources.get_distribution('numpy')
	except pkg_resources.DistributionNotFound:
		print "Numpy package not available."
		exit(0)
	else:
		import numpy as np
		
	try:
		pkg_resources.get_distribution('pandas')
	except pkg_resources.DistributionNotFound:
		print "Pandas package not available."
		exit(0)
	else:
		import pandas as pd
		import pandas.rpy.common as com
		
	try:
		pkg_resources.get_distribution('rpy2')
	except pkg_resources.DistributionNotFound:
		print "rpy2 package not available."
		exit(0)
	else:
		import rpy2.robjects as robjects
		
	
	print robjects.r("setwd('"+phjPathToNewWorkingDirectory+"')")
	
	print robjects.r.load(".RData")
	
	myRData = com.load_data(phjRData)
	
	return myRData
	
	

