from flask import Flask, render_template, request, redirect, session, jsonify
import pyexcel
import flask_excel as excel
app=Flask(__name__)

@app.route("/upload", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        workbook = request.get_records(field_name='file')
        list_value = []
        for item in workbook:
            list_key = []
            list_item_value = []
            for key,value in item.items():
                list_key.append(key)
                list_item_value.append(value)
            list_value.append(list_item_value)
        return excel.make_response_from_records(workbook,"xlsx", file_name="download")
    if request.method == "GET":
        return render_template("demo_test.html")

if __name__ == "__main__":
    excel.init_excel(app)
    app.run(debug=True)