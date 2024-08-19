import xlwings as xw
from requests import get
import datetime as dt

def get_ecb_rates():
    entrypoint = 'https://data-api.ecb.europa.eu/service/'
    resource = 'data'
    flowRef = 'FM'
    key = 'D.U2.EUR.4F.KR.DFR.LEV'

    parameters = {
        'startPeriod': (dt.datetime.today() - dt.timedelta(1000)).strftime('%Y-%m-%d'),
        'format': 'csvdata'
    }

    request_url = f"{entrypoint}{resource}/{flowRef}/{key}"
    response = get(request_url, params=parameters)
    csv_data = response.content.decode('utf-8')

    # Split CSV data into rows and columns
    data_rows = [line.split(',') for line in csv_data.splitlines()]

    # Find the maximum number of columns in any row
    max_columns = max(len(row) for row in data_rows)

    # Ensure all rows have the same number of columns by filling with empty strings
    uniform_data = [row + [''] * (max_columns - len(row)) for row in data_rows]

    return uniform_data

def main():
    uniform_data = get_ecb_rates()
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    sheet.range("A1").value = uniform_data
