#Importing Data Practice 
#Code reference: Ch.9 Data Wrangling with Python.
#Handcoded and commented for learning purposes (not pasted) 

# Python library for working with Excel files.
import xlrd

# Python library that allows us to look at basic features of our data.
import agate


#Our Excel file is located in our current directory.
excl_file = 'unicef_oct_2014.xls'

#Get data from the Excel file excl_file into variable workbook
workbook = xlrd.open_workbook(excl_file)

#Number of sheets in our workbook
workbook.nsheets

#Name's of every single sheet, in this case only one name.
workbook.sheet_names()

#The right hand of the assignment operator gets the first sheet from the list of sheets.
sheet = workbook.sheets()[0]

#Get the number of rows of our selected sheet.
sheet.nrows

#Selects the first row, in this case the title.
sheet.row_values(0)

#Iterate each row
#for r in range(sheet.nrows):
    #prints the row at row r of our sheet.
    #print (r, sheet.row(r))

#zip() maps the similar index of multiple 
#containers (in this case row 4 and 5) so that they 
#can be used just using as single entity.
title_rows = zip(sheet.row_values(4), sheet.row_values(5))
#title_rows

# Iterate each tuple in title_rows
# to turn the titles into a list of strings called titles
# t[0] is the first value of tuple t
# t[1] is the second value of tuple t
titles = [t[0] + ' ' + t[1] for t in title_rows]
#Our tuples are not needed no more because we have a list of strings.

# Our titles are now messy (there can be leading spaces like ' Female')
# Use string.strip() to remove leading/trailing spaces
titles = [t.strip() for t in titles]
#print(titles)

#*************************************
#Now our titles variable has a clean
#list of strings for use with the agate library
#*************************************

#Lets focus on country data
country_rows = [sheet.row_values(r) for r in range(6, 114)]
#print(country_rows)

# ctype_text is an xlrd built-in tool to help define columns
from xlrd.sheet import ctype_text
import agate

#We need to define the types of data lists titles and country_rows
# Using agate, we can figure out our data's types
text_type = agate.Text()
number_type = agate.Number()
boolean_type = agate.Boolean()
date_type = agate.Date()

example_row = sheet.row(6)
print('example_row:')
print(example_row)

# ctype returns the type of the value
# from the documentation:
# 0 = empty string
# 1 = string
# 2, 3 = float
# 4 = boolean (0 for false, 1 for true)
# 5 = Excel cell error
# 6 = blank
#In this case our first value in our row is a text, so ctype = 1 
print(example_row[0].ctype)
print(example_row[0].value)
print('ctype_text: ')

# Prints what I have in the the previous comment block. 
# Use as a visual reference
print(ctype_text)

#Now we need to make a list of types for our agate library.
types = []
#Iterate over our row and use ctype to match he column types.
for v in example_row:
    #Maps the integer v.ctype of each column with ctype_text dict.
    #to make them readable with a string.
    value_type = ctype_text[v.ctype]
    #check what the string holds.
    if value_type == 'text':
        types.append(text_type)
    elif value_type == 'number':
        types.append(number_type)
    elif value_type == 'xldate':
        types.append(date_type)
    else:
        #It is recommmended if there is no type match
        #append the text column type.
        types.append(text_type)
        
        
#import the results into our agate table:
#NOTE: Gives a cast error cannot convert '-' to decimal.
#table = agate.Table(country_rows, titles, types)

#Define a function to remove bad chars
# Its better to have None instead of bad chars.
def remove_bad_chars(val):
    if val == '-':
        return None
    return val

cleaned_rows = []
for row in country_rows:
    #For each column in our current row, remove bad chars.
    cleaned_row = [remove_bad_chars(rv) for rv in row]
    #Once cleaned, append it to our new list.
    cleaned_rows.append(cleaned_row)

    
# A reusable function for cleaning arrays with any operation/function we pass.
def get_new_array(old_array, function_to_clean):
    new_arr = []
    for row in old_array:
        cleaned_row = [function_to_clean(rv) for rv in row]
        new_arr.append(cleaned_row)
    return new_arr

#Remove bad chars from country_rows
cleaned_rows = get_new_array(country_rows, remove_bad_chars)
#print('cleaned_rows:')
#print(cleaned_rows)


#Finally, create the agate table:
table = agate.Table(cleaned_rows, titles, types)
#print('Result Agate Table:')
#print(table)

#Use this print method to see what the table looks like instead of the normal print method above.
#max_columns: The maximum number of columns to display before truncating the data.
table.print_table(max_columns=7)
