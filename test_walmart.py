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



#Iterate through each column
for col_num in range(1, first_sheet.max_column):
    #print(col_num)
    #Iterate through rows, skip first row
    for row_num in range(1, first_sheet.max_row):
        #print("Step 1")
        data = first_sheet.cell(row=row_num, column = col_num).value
        #print("Good")
        new_first_sheet.cell(row = row_num, column = col_num).value = data

#Pass file name to save
new_wb.save("test_copy.xlsx")
