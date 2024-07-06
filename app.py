from flask import Flask, render_template, request, send_file
import openpyxl
from openpyxl.styles import Font, numbers
from io import BytesIO

app = Flask(__name__)

@app.template_filter('format_number')
def format_number(value, currency_symbol):
    return f"{currency_symbol}{value:,.2f}"

@app.route('/', methods=['GET', 'POST'])
def fibonacci_calculator():
    if request.method == 'POST':
        start_number = float(request.form['start_number'])
        fibonacci_range = int(request.form['fibonacci_range'])
        currency_code = request.form['currency_code']
        currency_symbol = CURRENCY_SYMBOLS[currency_code]

        # Calculate the Fibonacci sequence
        a, b = 0, start_number
        fibonacci_values = []
        for i in range(fibonacci_range + 1):
            fibonacci_values.append(b)
            a, b = b, a + b

        return render_template('index.html', fibonacci_values=fibonacci_values, fibonacci_range=fibonacci_range, currency_symbol=currency_symbol, currency_symbols=CURRENCY_SYMBOLS)

    return render_template('index.html', currency_symbols=CURRENCY_SYMBOLS)

@app.route('/export', methods=['POST'])
def export_to_excel():
    start_number = float(request.form['start_number'])
    fibonacci_range = int(request.form['fibonacci_range'])
    currency_code = request.form['currency_code']
    currency_symbol = CURRENCY_SYMBOLS[currency_code]

    # Calculate the Fibonacci sequence
    a, b = 0, start_number
    fibonacci_values = []
    for i in range(fibonacci_range + 1):
        fibonacci_values.append(b)
        a, b = b, a + b

    # Save the results to an Excel file
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Fibonacci Results"

    sheet['A1'] = "Iteration"
    sheet['B1'] = "Value"
    sheet['A1'].font = Font(bold=True)
    sheet['B1'].font = Font(bold=True)

    for i, value in enumerate(fibonacci_values):
        sheet[f'A{i+2}'] = i
        sheet[f'B{i+2}'] = value
        sheet[f'B{i+2}'].number_format = f'"{currency_symbol}"#,##0.00'

    # Create an in-memory file-like object
    file_buffer = BytesIO()
    workbook.save(file_buffer)
    file_buffer.seek(0)

    return send_file(
        file_buffer,
        as_attachment=True,
        download_name="fibonacci_results.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

CURRENCY_SYMBOLS = {
    "USD": "$",
    "EUR": "€",
    "GBP": "£",
    "JPY": "¥",
    "AUD": "AU$",
    "CAD": "CA$",
    "CHF": "CHF",
    "CNY": "¥",
    "HKD": "HK$",
    "NZD": "NZ$",
    "SEK": "SEK",
    "KRW": "₩",
    "MXN": "MX$",
    "INR": "₹",
    "RUB": "₽",
    "BRL": "R$",
    "TRY": "₺",
    "ZAR": "R",
    "PLN": "zł",
    "NOK": "kr",
    "DKK": "kr",
    "HUF": "Ft",
    "CZK": "Kč",
    "AED": "AED",
    "SAR": "SAR",
    "THB": "฿",
    "ILS": "₪",
    "SGD": "S$",
    "PHP": "₱",
    "IDR": "Rp",
    "MYR": "RM",
    "VND": "₫",
    "EGP": "EGP",
    "CLP": "CLP",
    "ARS": "ARS",
    "PEN": "S/",
    "QAR": "QAR",
    "KWD": "KWD",
    "MAD": "MAD",
    "DZD": "DZD",
    "TND": "TND",
    "JOD": "JOD",
    "LKR": "LKR",
    "BHD": "BHD",
    "OMR": "OMR",
    "JMD": "J$",
    "TTD": "TTD",
    "KES": "KES",
    "UGX": "UGX",
    "GHS": "GHS",
    "ZMW": "ZMW",
    "NAD": "NAD",
    "BWP": "BWP"
}

if __name__ == '__main__':
    app.run(debug=True)