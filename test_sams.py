import openpyxl
import os

#MAC filepath
filepath = os.path.join('/Users', 'KyleShare', 'Programming', 'caravan', 'SAMS.XLSX' )

#WINDOWS filepath
#filepath = os.path.join('C:\\', 'Users', 'CaravanArms', 'Desktop', 'SAMS.XLSX' )

#Get workbook from filepath
wb = openpyxl.load_workbook(filepath)

#Get all sheetnames
sheet_names = wb.sheetnames

#Get first worksheet, since wb.active is set to 0 by default
first_sheet = wb.active

#Create new workbook
new_wb = openpyxl.Workbook()
new_first_sheet = new_wb.active

#first_po = Get PO
#first_po_data...
#Copy pt 1 relevant info to WRITING ROW
#Next reading line
#Copy pt 2 of relevant info to WRITING ROW
#increase reading row
#increase writing row
#if current_po == first_po
  #copy pt 1 relevant info from first_po or previous po to WRITING ROW
  #copy pt 2 of relevant info to WRITING ROW
#else:
  #first_po = Get PO



def titles():
    titles = ["ACCOUNT(SBT CODE)", "PO#", "PO LINE", "CUSTOMER NAME", "ADDRESS 1(2ND LINE)", \
    "PHONE# (3RD LINE)", "ADDRESS 2", "CARRIER", "ITEM#", "ITEM DESCRIPTION", \
    "UNIT PRICE", "QTY", "LINE TOTAL", "TERMS"]
    title_index = 0
    for column_num in range(1, 15):
        new_first_sheet.cell(row = 1, column = column_num).value = titles[title_index]
        title_index += 1

def record_type():
       counter = 0
       first_po = None
       #If new_line is true, create new line
       new_line = False
       for row_num in range(2, first_sheet.max_row + 1):
           #If PO does not change, copy details of order to same header
           if first_sheet.cell(row = row_num, column = 1).value == first_po:
               #If new_line == True, we want a new line even though po stays same
               #Therefore, don't increase counter
               if new_line == False:
                   print("new_line is False")
                   new_line = True
                   counter += 1
                   print("counter is", counter)
               print("These functions will write to line", row_num - counter)
               po_line(writing_row = row_num - counter, row = row_num)
               item_num(writing_row = row_num - counter, row = row_num)
               item_desc(writing_row = row_num - counter, row = row_num)
               unit_price(writing_row = row_num - counter, row = row_num)
               quantity(writing_row = row_num - counter, row = row_num)
               continue

           #If PO changes, copy header of order
           new_line = False
           first_po = first_sheet.cell(row = row_num, column = 1).value
           account(writing_row = row_num - counter, row = row_num)
           po_num(writing_row = row_num - counter, row = row_num)
           customer_name(writing_row = row_num - counter, row = row_num)
           address_1(writing_row = row_num - counter, row = row_num)
           phone_num(writing_row = row_num - counter, row = row_num)
           address_2(writing_row = row_num - counter, row = row_num)
           carrier(writing_row = row_num - counter, row = row_num)
           terms(writing_row = row_num - counter, row = row_num)


def account(writing_row, row):
    new_first_sheet.cell(row = writing_row, column = 1).value = 'WMECOM'

def po_num(writing_row, row):
    po = first_sheet.cell(row = row, column = 1).value
    new_first_sheet.cell(row = writing_row, column = 2).value = po

def po_line(writing_row, row):
    line = first_sheet.cell(row = row, column = 13).value
    new_first_sheet.cell(row = writing_row, column = 3).value = line


def customer_name(writing_row, row):
    name = first_sheet.cell(row = row, column = 63).value
    new_first_sheet.cell(row = writing_row, column = 4).value = name

#<Street address, Appt/Suite>
#Appt/Suite may be in same cell as street address or 1 column right
def address_1(writing_row, row):
    street_address = first_sheet.cell(row = row, column = 64).value
    apartment = first_sheet.cell(row = row, column = 65).value

    #if apartment exists on next column, add it to street address
    if apartment != None:
      street_address = "{} {}".format(street_address, apartment)

    new_first_sheet.cell(row = writing_row, column = 5).value = street_address

def phone_num(writing_row, row):
    phone = first_sheet.cell(row = row, column = 78).value
    new_first_sheet.cell(row = writing_row, column = 6).value = phone

#<City, State Zip>
def address_2(writing_row, row):
    city = first_sheet.cell(row = row, column = 66).value
    state = first_sheet.cell(row = row, column = 67).value
    zip_code = first_sheet.cell(row = row, column = 68).value

    address2 = "{}, {} {}".format(city, state, zip_code)
    new_first_sheet.cell(row = writing_row, column = 7).value = address2

def carrier(writing_row, row):
    new_first_sheet.cell(row = writing_row, column = 8).value = '3PT FDXG'

def item_num(writing_row, row):
    item_num = first_sheet.cell(row = row, column = 19).value
    new_first_sheet.cell(row = writing_row, column = 9).value = item_num

def item_desc(writing_row, row):
    item_desc = first_sheet.cell(row = row, column = 21).value
    new_first_sheet.cell(row = writing_row, column = 10).value = item_desc

def unit_price(writing_row, row):
    unit_price = first_sheet.cell(row = row, column = 16).value
    new_first_sheet.cell(row = writing_row, column = 11).value = unit_price

def quantity(writing_row, row):
    quantity = first_sheet.cell(row = row, column = 14).value
    new_first_sheet.cell(row = writing_row, column = 12).value = quantity


def terms(writing_row, row):
    new_first_sheet.cell(row = writing_row, column = 14).value = 'NET 60'

def fill_empty_cells():
    for row_num in range(2, new_first_sheet.max_row + 1):
        #If cell is empty, set account, poline, customer, address1, phone, address2, carrier, terms == to previous
        if (new_first_sheet.cell(row = row_num, column = 1).value) == None:

            previous_acc = new_first_sheet.cell(row = row_num - 1, column = 1).value
            new_first_sheet.cell(row = row_num, column = 1).value = previous_acc

            previous_po_num = new_first_sheet.cell(row = row_num - 1, column = 2).value
            new_first_sheet.cell(row = row_num, column = 2).value = previous_po_num

            previous_customer = new_first_sheet.cell(row = row_num - 1, column = 4).value
            new_first_sheet.cell(row = row_num, column = 4).value = previous_customer

            previous_address1 = new_first_sheet.cell(row = row_num - 1, column = 5).value
            new_first_sheet.cell(row = row_num, column = 5).value = previous_address1

            previous_phone = new_first_sheet.cell(row = row_num - 1, column = 6).value
            new_first_sheet.cell(row = row_num, column = 6).value = previous_phone

            previous_address2 = new_first_sheet.cell(row = row_num - 1, column = 7).value
            new_first_sheet.cell(row = row_num, column = 7).value = previous_address2

            previous_carrier = new_first_sheet.cell(row = row_num - 1, column = 8).value
            new_first_sheet.cell(row = row_num, column = 8).value = previous_carrier

            previous_terms = new_first_sheet.cell(row = row_num - 1, column = 14).value
            new_first_sheet.cell(row = row_num, column = 14).value = previous_terms

#Use quantity and Unit price to calculate line total
def line_total():
    for row_num in range(2, new_first_sheet.max_row + 1):
      line_total = new_first_sheet.cell(row = row_num, column = 11).value * \
      new_first_sheet.cell(row = row_num, column = 12).value
      new_first_sheet.cell(row = row_num, column = 13).value = line_total

def main():
    titles()
    record_type()
    fill_empty_cells()
    line_total()


main()

new_wb.save("SBT_SAMS.xlsx")