from flask import Flask, request, send_file
import xlwings as xw
import tempfile
import json

app = Flask(__name__)

@app.route('/update-xlsm', methods=['POST'])
def update_xlsm():
    file = request.files['file']
    metadata = json.loads(request.form['metadata'])

    with tempfile.NamedTemporaryFile(suffix=".xlsm", delete=False) as tmp:
        file.save(tmp.name)
        wb = xw.Book(tmp.name)
        sht = wb.sheets[0]

        for row in range(4, 100):  # Assume contract in column C starting at row 4
            if sht.range(f'C{row}').value == metadata["contract_number"]:
                for col in range(1, 30):  # Assume month headers in row 3
                    month_value = sht.range((3, col)).value
                    if month_value and month_value.lower() == metadata["month"].lower():
                        sht.range((row, col)).value = metadata["production_kWh"]
                        break
                break

        wb.save()
        wb.close()
        return send_file(tmp.name, as_attachment=True)
