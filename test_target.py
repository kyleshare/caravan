import openpyxl
import os

#MAC filepath
filepath = os.path.join('/Users', 'KyleShare', 'Programming', 'caravan', 'TRAGET.XLSX' )

#WINDOWS filepath
#filepath = os.path.join('C:\\', 'Users', 'CaravanArms', 'Desktop', 'TARGET.XLSX' )

#Get workbook from filepath
wb = openpyxl.load_workbook(filepath)

#Get all sheetnames
sheet_names = wb.sheetnames

#Get first worksheet, since wb.active is set to 0 by default
first_sheet = wb.active

#Create new workbook
new_wb = openpyxl.Workbook()
new_first_sheet = new_wb.active

#Create final workbook
final_wb = openpyxl.Workbook()
final_first_sheet = final_wb.active


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
        new_first_sheet.cell(row = row_num, column = 1).value = 'TARGET'

#Initial 0's omitted on downloads, make PO 10 digits
def po_num(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        po = first_sheet.cell(row = row_num, column = column_num).value
        po = str(po)
        po = po.zfill(10)
        new_first_sheet.cell(row = row_num, column = 2).value = po
#2
def po_line(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        line = first_sheet.cell(row = row_num, column = column_num).value
        new_first_sheet.cell(row = row_num, column = 3).value = line

#1
def customer_name(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        name = first_sheet.cell(row = row_num, column = column_num).value
        new_first_sheet.cell(row = row_num, column = 4).value = name

#1, bill or ship?
def address_1(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        address = first_sheet.cell(row = row_num, column = column_num).value
        new_first_sheet.cell(row = row_num, column = 5).value = address

def phone_num(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        phone = first_sheet.cell(row = row_num, column = column_num).value
        new_first_sheet.cell(row = row_num, column = 6).value = phone

#<City, State Zip>
#Initial 0's omitted on downloads, make Zip Code 5 digits
def address_2(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        #Assumes city state zip are next to each other
        city = first_sheet.cell(row = row_num, column = column_num).value
        state = first_sheet.cell(row = row_num, column = column_num + 1).value
        zip_code = first_sheet.cell(row = row_num, column = column_num + 2).value
        zip_code = str(zip_code)
        zip_code = zip_code.zfill(5)

        address2 = "{}, {} {}".format(city, state, zip_code)
        new_first_sheet.cell(row = row_num, column = 7).value = address2

def carrier(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        carrier = first_sheet.cell(row = row_num, column = 24).value
        new_first_sheet.cell(row = row_num, column = 8).value = carrier

def item_num(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        item_num = first_sheet.cell(row = row_num, column = 18).value
        new_first_sheet.cell(row = row_num, column = 9).value = item_num

#2
def item_desc(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        item_desc = first_sheet.cell(row = row_num, column = 20).value
        new_first_sheet.cell(row = row_num, column = 10).value = item_desc

#2
def unit_price(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        price = first_sheet.cell(row = row_num, column = 15).value
        new_first_sheet.cell(row = row_num, column = 11).value = price
#2
def quantity(column_num):
    for row_num in range(2, first_sheet.max_row + 1):
        quantity = first_sheet.cell(row = row_num, column = 13).value
        new_first_sheet.cell(row = row_num, column = 12).value = quantity

def terms():
    for row_num in range(2, first_sheet.max_row + 1):
        new_first_sheet.cell(row = row_num, column = 14).value = 'NET 30'

def count_rows():
    total_rows = 0
    #Iterate through new sheet
    for row_num in range(2, new_first_sheet.max_row + 1):
        #Count PO Lines to determine # of rows needed
        if new_first_sheet.cell(row = row_num, column = 3).value:
            total_rows += 1
    return total_rows
    
def remove_empty_cells(total_rows):
    #Iterate through each column
    for column_num in range(1, first_sheet.max_column + 1):
        #For each column, row initialized to 1
        final_sheet_row = 1
        #Iterate through each row in a column
        for row_num in range(1, new_first_sheet.max_row + 1):
                new_cell = new_first_sheet.cell(row = row_num, column = column_num).value
                #If cell is not empty and row is < total rows
                if new_cell not in [None, "None, None None"] and final_sheet_row < total_rows:
                    final_first_sheet.cell(row = final_sheet_row, column = column_num).value = new_cell
                    final_sheet_row += 1
    return final_wb

#Use quantity and Unit price to calculate line total
def line_total():
    for row_num in range(2, final_first_sheet.max_row + 1):
      line_total = final_first_sheet.cell(row = row_num, column = 11).value * \
      final_first_sheet.cell(row = row_num, column = 12).value
      final_first_sheet.cell(row = row_num, column = 13).value = line_total



def main():
    titles()
    account()
    po_num(1)
    po_line(12)
    customer_name(62)
    address_1(63)
    phone_num(52)
    address_2(65)
    carrier(24)
    item_num(18)
    item_desc(20)
    unit_price(15)
    quantity(13)
    terms()

    total_rows = count_rows()
    remove_empty_cells(total_rows)

    line_total()

main()

#Pass file name to save, only save final file
final_wb.save("SBT_TARGET.xlsx")