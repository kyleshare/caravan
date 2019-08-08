#import json
import openpyxl
import os

#If you have a Python object, you can convert it into a JSON string by using the json.dumps() method.
#obj = json.dumps()

#directory = os.getcwd()

#function takes file/folder names in path, returns correct file path based on os
filepath = os.path.join('/Users', 'KyleShare', 'Programming', 'caravan', '1WALMART W.H.XLSX' )
wb = openpyxl.load_workbook(filepath)
print(type(wb))


#Gets all sheetnames
sheet_names = wb.sheetnames
print("All sheet_names are", sheet_names)
#Gets first worksheet, since wb.active is set to 0 by default
first_sheet = wb.active
print("The first sheet is", first_sheet)

#Create new workbook
new_wb = openpyxl.Workbook()
new_first_sheet = new_wb.active
print("The first sheet is", new_first_sheet)
#print(new_wb.sheetnames)

def account():
    for row_num in range(2, first_sheet.max_row):
        new_first_sheet.cell(row = row_num, column = 1).value = 'wmecom'

def po_num():
    for row_num in range(2, first_sheet.max_row):
        po = first_sheet.cell(row = row_num, column = 1).value
        new_first_sheet.cell(row = row_num, column = 2).value = po

def po_line():
    for row_num in range(2, first_sheet.max_row):
        line = first_sheet.cell(row = row_num, column = 16).value
        new_first_sheet.cell(row = row_num, column = 3).value = line

#Maybe I should hide this for legal reasons?
def customer_name():
    for row_num in range(2, first_sheet.max_row):
        name = first_sheet.cell(row = row_num, column = 5).value
        new_first_sheet.cell(row = row_num, column = 4).value = name

def address_1():
    for row_num in range(2, first_sheet.max_row):
        address = first_sheet.cell(row = row_num, column = 6).value
        new_first_sheet.cell(row = row_num, column = 5).value = address

#This might be wrong. Im copying from shipping address but maybe it should 
def phone_num():
    for row_num in range(2, first_sheet.max_row):
        phone = first_sheet.cell(row = row_num, column = 7).value
        new_first_sheet.cell(row = row_num, column = 6).value = phone


def address_2():
    for row_num in range(2, first_sheet.max_row):
        address2 = first_sheet.cell(row = row_num, column = 10).value
        new_first_sheet.cell(row = row_num, column = 7).value = address2

def carrier():
    pass

def item_num():
    pass


'''
#Iterate through each column
for col_num in range(1, first_sheet.max_column):
    #print(col_num)
    #Iterate through rows, skip first row
    for row_num in range(1, first_sheet.max_row):
        #print("Step 1")
        data = first_sheet.cell(row=row_num, column = col_num).value
        #print("Good")
        new_first_sheet.cell(row = row_num, column = col_num).value = data
'''
account()
po_num()
po_line()
customer_name()
address_1()
phone_num()
address_2()

#Pass file name to save
new_wb.save("test_copy.xlsx")
