from flask import Flask, render_template
import pandas as pd
import os

app = Flask(__name__)

EXCEL_FILE = r'C:\Users\msi\PycharmProjects\PythonProject7\تقدم العمل.xlsx'

@app.route('/')
def index():
    if not os.path.exists(EXCEL_FILE):
        return "ملف Excel غير موجود"

    xls = pd.ExcelFile(EXCEL_FILE)
    sheet_names = xls.sheet_names

    sheets_data = []

    for sheet in sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.fillna("", inplace=True)
        rows = df.values.tolist()[3:]
        headers = df.columns.tolist()
        sheets_data.append({
            "name": sheet,
            "headers": headers,
            "rows": rows
        })

    return render_template('index.html', sheets_data=sheets_data)

if __name__ == '__main__':
    app.run(debug=True)