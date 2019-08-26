#Kyle Share
import openpyxl
import os

#MAC filepath
#filepath = os.path.join('/Users', 'KyleShare', 'Programming', 'caravan', '1WALMART W.H.XLSX' )
#description_path = os.path.join('/Users', 'KyleShare', 'Programming', 'caravan', 'Item Description.xlsx')

#WINDOWS filepath
filepath = os.path.join('C:\\', 'Users', 'CaravanArms', 'Desktop', 'WALMART W.H.XLSX' )
description_path = os.path.join('C:\\', 'Users', 'AKim', 'Desktop', 'Item Description.xlsx.')

#Get workbook from filepath
wb = openpyxl.load_workbook(filepath)
description_wb = openpyxl.load_workbook(description_path)

#Get all sheetnames
sheet_names = wb.sheetnames
description_sheet_names = description_wb.sheetnames

#Get first worksheet, since wb.active is set to 0 by default
first_sheet = wb.active
description_sheet = description_wb.active

#Create new workbook
new_wb = openpyxl.Workbook()
new_first_sheet = new_wb.active

#Create dictionary of (ID: item description) pairs
def description_dict():
    description_list = []
    #Iterate through description excel sheet
    for row_num in range(2, description_sheet.max_row):
        key = str((description_sheet.cell(row=row_num, column=1).value))
        description = (description_sheet.cell(row=row_num, column=2).value)
        #Create tuples, append them to description list
        tupl = (key, description)
        description_list.append(tupl)
        #Make description dictionary global so item desctiption function can access it
        global description_dict
        description_dict = dict(description_list)

def price_dict():
    price_list = []
    #Iterate through prices excel sheet
    for row_num in range(2, description_sheet.max_row):
        key = str((description_sheet.cell(row=row_num, column=1).value))
        #Walmart price is column 3
        price = (description_sheet.cell(row=row_num, column=3).value)
        #Create tuples, append them to description list
        tupl = (key, price)
        price_list.append(tupl)
        #Make description dictionary global so item desctiption function can access it
        global price_dict
        price_dict = dict(price_list)


def titles():
    titles = ["ACCOUNT(SBT CODE)", "PO#", "PO LINE", "CUSTOMER NAME", "ADDRESS 1(2ND LINE)", \
    "PHONE# (3RD LINE)", "ADDRESS 2", "CARRIER", "ITEM#", "ITEM DESCRIPTION", \
    "UNIT PRICE", "QTY", "LINE TOTAL", "TERMS"]
    title_index = 0
    for column_num in range(1, 15):
        new_first_sheet.cell(row = 1, column = column_num).value = titles[title_index]
        title_index += 1

def account():
    for row_num in range(2, first_sheet.max_row + 1):
        new_first_sheet.cell(row = row_num, column = 1).value = 'WMECOM'

#Initial 0's omitted on downloads, make PO 10 digits
def po_num():
    for row_num in range(2, first_sheet.max_row + 1):
        po = first_sheet.cell(row = row_num, column = 1).value
        po = str(po)
        po = po.zfill(10)
        new_first_sheet.cell(row = row_num, column = 2).value = po

def po_line():
    for row_num in range(2, first_sheet.max_row + 1):
        line = first_sheet.cell(row = row_num, column = 16).value
        new_first_sheet.cell(row = row_num, column = 3).value = line

#Maybe I should hide this for legal reasons?
def customer_name():
    for row_num in range(2, first_sheet.max_row + 1):
        name = first_sheet.cell(row = row_num, column = 5).value
        new_first_sheet.cell(row = row_num, column = 4).value = name

def address_1():
    for row_num in range(2, first_sheet.max_row + 1):
        address = first_sheet.cell(row = row_num, column = 9).value
        new_first_sheet.cell(row = row_num, column = 5).value = address

def phone_num():
    for row_num in range(2, first_sheet.max_row + 1):
        phone = first_sheet.cell(row = row_num, column = 7).value
        new_first_sheet.cell(row = row_num, column = 6).value = phone

#<City, State Zip>
#Initial 0's omitted on downloads, make Zip Code 5 digits
def address_2():
    for row_num in range(2, first_sheet.max_row + 1):
        city = first_sheet.cell(row = row_num, column = 11).value
        state = first_sheet.cell(row = row_num, column = 12).value
        zip_code = first_sheet.cell(row = row_num, column = 13).value
        zip_code = str(zip_code)
        zip_code = zip_code.zfill(5)


        address2 = "{}, {} {}".format(city, state, zip_code)
        new_first_sheet.cell(row = row_num, column = 7).value = address2

def carrier():
    for row_num in range(2, first_sheet.max_row + 1):
        carrier = first_sheet.cell(row = row_num, column = 24).value
        new_first_sheet.cell(row = row_num, column = 8).value = carrier

def item_num():
    for row_num in range(2, first_sheet.max_row + 1):
        item_num = first_sheet.cell(row = row_num, column = 18).value
        new_first_sheet.cell(row = row_num, column = 9).value = item_num

#May have 2 item desc
def item_desc():
    for row_num in range(2, first_sheet.max_row + 1):
        item_num = first_sheet.cell(row = row_num, column = 18).value
        item_desc = description_dict[item_num]
        new_first_sheet.cell(row = row_num, column = 10).value = item_desc

def unit_price():
    for row_num in range(2, first_sheet.max_row + 1):
        item_num = first_sheet.cell(row = row_num, column = 18).value
        item_price = price_dict[item_num]
        new_first_sheet.cell(row = row_num, column = 11).value = item_price

#Walmart stores quantity as text, convert to int
def quantity():
    for row_num in range(2, first_sheet.max_row + 1):
        quantity = first_sheet.cell(row = row_num, column = 21).value
        quantity = int(quantity)
        new_first_sheet.cell(row = row_num, column = 12).value = quantity


def terms():
    for row_num in range(2, first_sheet.max_row + 1):
        new_first_sheet.cell(row = row_num, column = 14).value = 'NET 30'

#Use quantity and Unit price to calculate line total
def line_total():
    for row_num in range(2, new_first_sheet.max_row + 1):
            line_total = new_first_sheet.cell(row = row_num, column = 11).value * \
            new_first_sheet.cell(row = row_num, column = 12).value
            new_first_sheet.cell(row = row_num, column = 13).value = line_total

def main():
    description_dict()
    price_dict()
    titles()
    account()
    po_num()
    po_line()
    customer_name()
    address_1()
    phone_num()
    address_2()
    carrier()
    item_num()
    item_desc()
    unit_price()
    quantity()
    terms()
    #Implement after unit price is fixed
    line_total()

main()

#Pass file name to save
new_wb.save("SBT_WALMART.xlsx")
