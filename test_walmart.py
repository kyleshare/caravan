import openpyxl
import os

#MAC filepath
filepath = os.path.join('/Users', 'KyleShare', 'Programming', 'caravan', '1WALMART W.H.XLSX' )

#WINDOWS filepath
#filepath = os.path.join('C:', 'Users', 'CaravanArms', 'Desktop', 'WALMART W.H.XLSX' )


#Get workbook from filepath
wb = openpyxl.load_workbook(filepath)

#Get all sheetnames
sheet_names = wb.sheetnames

#Get first worksheet, since wb.active is set to 0 by default
first_sheet = wb.active

#Create new workbook
new_wb = openpyxl.Workbook()
new_first_sheet = new_wb.active


def titles():
        titles = ["ACCOUNT(SBT CODE)", "PO#", "PO LINE", "CUSTOMER NAME", "ADDRESS 1(2ND LINE)", \
        "PHONE# (3RD LINE)", "ADDRESS 2", "CARRIER", "ITEM#", "ITEM DESCRIPTION", "QTY", \
        "UNIT PRICE", "TERMS"]
        title_index = 0
        for column_num in range(1, 14):
            new_first_sheet.cell(row = 1, column = column_num).value = titles[title_index]
            title_index += 1

def account():
    for row_num in range(2, first_sheet.max_row + 1):
        new_first_sheet.cell(row = row_num, column = 1).value = 'WMECOM'

def po_num():
    for row_num in range(2, first_sheet.max_row + 1):
        po = first_sheet.cell(row = row_num, column = 1).value
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

def address_2():
    for row_num in range(2, first_sheet.max_row + 1):
        city = first_sheet.cell(row = row_num, column = 11).value
        state = first_sheet.cell(row = row_num, column = 12).value
        zip_code = first_sheet.cell(row = row_num, column = 13).value

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

def item_desc():
    for row_num in range(2, first_sheet.max_row + 1):
        item_desc = first_sheet.cell(row = row_num, column = 20).value
        new_first_sheet.cell(row = row_num, column = 10).value = item_desc

def quantity():
    for row_num in range(2, first_sheet.max_row + 1):
        quantity = first_sheet.cell(row = row_num, column = 21).value
        new_first_sheet.cell(row = row_num, column = 11).value = quantity

def unit_price():
    pass

def terms():
    for row_num in range(2, first_sheet.max_row + 1):
        new_first_sheet.cell(row = row_num, column = 13).value = 'NET 30'

def main():
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
    quantity()
    unit_price()
    terms()

main()

#Pass file name to save
new_wb.save("test_copy.xlsx")
