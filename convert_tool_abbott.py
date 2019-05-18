from flask import Flask, render_template, request, redirect, session, jsonify
import flask_excel as excel 
import xlrd
from collections import OrderedDict
from datetime import datetime

abivin_workbook = xlrd.open_workbook("Import_Order_Template.xlsx")
abivin_sheet = abivin_workbook.sheet_by_index(0)
key_order = abivin_sheet.row_values(0)
expired_date_int = abivin_sheet.row_values(1)[17]
expired_date = datetime.fromordinal(datetime(1900, 1, 1).toordinal() + int(expired_date_int) - 2)
app = Flask(__name__)

@app.route('/convert_tool_abbott', methods = ["GET","POST"])
def convert_tool_abbott():
  if request.method == "GET":
    return render_template('abbott.html')
  if request.method == "POST":
    abbott_workbook = request.get_records(field_name = "file")
    converted_list = []
    for item in abbott_workbook:
      dictionary = OrderedDict(
        {
          key_order[0] : item["Ngày đặt hàng"],
          key_order[1] : item["Số đơn hàng"],
          key_order[2] : "SALES",
          key_order[3] : item["Mã KH"],
          key_order[4] : item["Mã SP"],
          key_order[5] : "7:30-17:00",
          key_order[6] : item["Số lượng (thùng)"],
          key_order[7] : item["Số lượng (lẻ)"],
          key_order[8] : item["Thành tiền (VND)"],
          key_order[9] : 0,
          key_order[10] : 0,
          key_order[11] : item["Số tiền chiết khấu"],
          key_order[12] : 0,
          key_order[13] : 0,
          key_order[14] : 5,
          key_order[15] : "FALSE",
          key_order[16] : "ABBOTT",
          key_order[17] : expired_date,
        }
      )
      converted_list.append(dictionary)
    return excel.make_response_from_records(converted_list, "xlsx", file_name= "order_abbott")

if __name__ == '__main__':
  excel.init_excel(app)
  app.run(debug=True)
 