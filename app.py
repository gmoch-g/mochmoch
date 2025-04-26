from flask import Flask, render_template, request, jsonify
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import datetime
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_FILE_PATH = os.path.join(BASE_DIR, 'data.xlsx')

app = Flask(__name__)

# ألوان التنسيقات
green = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # أخضر
white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")  # أبيض
orange = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")  # برتقالي فاتح
gray = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")  # رصاصي فاتح
red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")    # أحمر
yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")  # أصفر

# ========== تنسيقات حسب كل ورقة ==========

def style_planning_projects(sheet, row):
    cell_i = sheet[f"I{row}"]
    if cell_i.value and "تم الاعلان" in str(cell_i.value).strip():
        cell_i.fill = green
        sheet[f"H{row}"].fill = white

    prev_fill = None
    for col in "JKLMNOPQRSTU":
        cell = sheet[f"{col}{row}"]
        if cell.value:
            if not prev_fill:
                cell.fill = green
                prev_fill = "green"
            elif prev_fill == "green":
                cell.fill = orange
                prev_fill = "orange"
            elif prev_fill == "orange":
                cell.fill = gray
                prev_fill = "gray"

def style_execution_projects(sheet, row):
    cell_j = sheet[f"J{row}"]
    try:
        value = float(cell_j.value)
        if value < 0:
            cell_j.fill = red
        elif value == 0:
            cell_j.fill = white
        else:
            cell_j.fill = green
    except:
        pass

def style_change_orders(sheet, row):
    today = datetime.today()
    columns = ["I", "K", "M", "P", "Q", "R", "U"]
    for col in columns:
        try:
            current = sheet[f"{col}{row}"]
            previous = sheet[f"{chr(ord(col) - 1)}{row}"]
            next_cell = sheet[f"{chr(ord(col) + 1)}{row}"]

            current_date = datetime.strptime(str(current.value).strip(), "%Y-%m-%d")
            prev_date = datetime.strptime(str(previous.value).strip(), "%Y-%m-%d")

            if next_cell.value:
                current.fill = green
            elif current_date > prev_date:
                current.fill = yellow
            elif current_date < today:
                current.fill = red
        except:
            pass

def style_extensions(sheet, row):
    today = datetime.today()
    for col in ["G", "J"]:
        try:
            current = sheet[f"{col}{row}"]
            previous = sheet[f"{chr(ord(col) - 1)}{row}"]
            next_cell = sheet[f"{chr(ord(col) + 1)}{row}"]

            current_date = datetime.strptime(str(current.value).strip(), "%Y-%m-%d")
            prev_date = datetime.strptime(str(previous.value).strip(), "%Y-%m-%d")

            if next_cell.value:
                current.fill = green
            elif current_date > prev_date:
                current.fill = yellow
            elif current_date < today:
                current.fill = red
        except:
            pass

def style_stops(sheet, row):
    today = datetime.today()
    col = "G"
    try:
        current = sheet[f"{col}{row}"]
        previous = sheet[f"{chr(ord(col) - 1)}{row}"]
        next_cell = sheet[f"{chr(ord(col) + 1)}{row}"]

        current_date = datetime.strptime(str(current.value).strip(), "%Y-%m-%d")
        prev_date = datetime.strptime(str(previous.value).strip(), "%Y-%m-%d")

        if next_cell.value:
            current.fill = green
        elif current_date > prev_date:
            current.fill = yellow
        elif current_date < today:
            current.fill = red
    except:
        pass

# تنسيق شامل بناءً على اسم الورقة
def apply_styles(sheet, row):
    title = sheet.title.strip()

    if title == "متابعة المشاريع وزارة التخطيط":
        style_planning_projects(sheet, row)
    elif title == "متابعة مشاريع قيد التنفيذ":
        style_execution_projects(sheet, row)
    elif title == "متابعة أوامر غيار 2025":
        style_change_orders(sheet, row)
    elif title == "متابعة المدد الاضافية":
        style_extensions(sheet, row)
    elif title == "متابعة التوقفات":
        style_stops(sheet, row)

# ========== عرض الواجهة الرئيسية ==========
@app.route('/')
def index():
    wb = load_workbook(EXCEL_FILE_PATH)
    sheets_data = []

    for sheet in wb.sheetnames:
        worksheet = wb[sheet]
        headers = [cell.value for cell in worksheet[1]]
        rows = []

        for row in worksheet.iter_rows(min_row=2):
            row_data = []
            for cell in row:
                cell_value = cell.value
                bg_color = None
                if cell.fill and cell.fill.fill_type == 'solid' and cell.fill.start_color:
                    start_color = cell.fill.start_color
                    rgb = getattr(start_color, 'rgb', None)
                    if rgb and isinstance(rgb, str) and len(rgb) == 8:
                        bg_color = rgb[2:]  # حذف alpha (أول خانتين)
                row_data.append({'value': cell_value, 'color': bg_color})
            rows.append(row_data)

        sheets_data.append({
            "name": sheet,
            "headers": headers,
            "rows": rows
        })

    return render_template('index.html', sheets_data=sheets_data)

# ========== حفظ التعديلات ==========
@app.route('/save', methods=['POST'])
def save():
    try:
        data = request.json
        wb = load_workbook(EXCEL_FILE_PATH)

        for sheet_name, rows in data.items():
            sheet = wb[sheet_name]
            for row_index, row in enumerate(rows[1:], start=2):  # تخطي الرؤوس
                for col_index, value in enumerate(row, start=1):
                    sheet.cell(row=row_index, column=col_index, value=value)
                # بعد تحديث الصف، نطبق التلوين المناسب
                apply_styles(sheet, row_index)

        wb.save(EXCEL_FILE_PATH)
        return jsonify({'status': 'success'})

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

# ========== تشغيل التطبيق ==========
if __name__ == '__main__':
    app.run(debug=True)
