import pyexcel
import xlrd
from collections import OrderedDict
from datetime import datetime
  
# CRUD an excel file + save as new file

workbook = xlrd.open_workbook("Import_Order_Template.xlsx")
sheet = workbook.sheet_by_index(0)
subject = sheet.row_values(0)
list_dictionary = []
for row_index in range(1,sheet.nrows):
  content = sheet.row_values(row_index)
  order_date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(content[0]) - 2)
  content[0] = order_date
  expire_date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(content[17]) - 2)
  content[17] = expire_date
  dictionary = OrderedDict()
  for index,item in enumerate(subject):
    dictionary[item] = content[index]
  list_dictionary.append(dictionary)
pyexcel.save_as(records = list_dictionary, dest_file_name = "result.xlsx")


# workbook = xlrd.open_workbook("out_put.xlsx")
# sheet = workbook.sheet_by_index(0)
# subject = sheet.row_values(0)
# list_dictionary = []
# for row_index in range(1,sheet.nrows):
#   content = sheet.row_values(row_index)
#   dictionary = {}
#   for index,item in enumerate(subject):
#     dictionary[item] = content[index]
#   list_dictionary.append(dictionary)
# print(list_dictionary)

# list1 = ["name", "address", "sale"]
# list2 = ["minh", "kim ma", 90]
# dictionary = {}
# menu = []
# for index, item in enumerate(list1):
#   dictionary[item] = list2[index]
# menu.append(dictionary)

# pyexcel.save_as(records = menu, dest_file_name = "test.xls")

