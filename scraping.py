import requests
from bs4 import BeautifulSoup
import pandas as pd


url = "https://bvb.ro/FinancialInstruments/Indices/IndicesProfiles.aspx?i=BET"

page = requests.get(url)
soup = BeautifulSoup(page.text, "html.parser")


table = soup.find('table', attrs={"class": "table table-hover dataTable no-footer generic-table compact"})

rows = table.find_all("tr")

headers = [header.text.strip() for header in rows[0].find_all('th')]

def convert_to_float(value):
    try:
        clean_value = value.replace(',', '.')
        return round(float(clean_value), 3)
    except ValueError:
        return value

def process_other(value):
    return value

data = []
for row in rows[1:]:
    cells = row.find_all('td')
    row_data = []
    for i, cell in enumerate(cells):
        cell_text = cell.text.strip()
        if 3 <= i <= 7:
            row_data.append(convert_to_float(cell_text))
        else:
            row_data.append(process_other(cell_text))
    data.append(row_data)

lung=len(data)
ok=0
row_data=[]
for i in range(lung):
    url="https://bvb.ro/FinancialInstruments/Details/FinancialInstrumentsDetails.aspx?s="+data[i][0]
    page = requests.get(url)
    soup = BeautifulSoup(page.text, "html.parser")
    table = soup.find('table', attrs={"id": "ctl00_body_ctl02_PricesControl_dvCPrices"})
    rows = table.find_all("tr")
    row_data=[]
    for row in rows[6:12]:
        cells = row.find_all('td')
        for j, cell in enumerate(cells):
            cell_text = cell.text.strip()
            if ok == 0 and j%2 == 0:
                headers.append(cell_text)
            else:
                if j%2 != 0:
                    row_data.append(convert_to_float(cell_text))
    data[i].extend(row_data)
    ok+=1

df = pd.DataFrame(data, columns=headers)

with pd.ExcelWriter('ConstituentiBet.xlsx', engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False,startrow=2, sheet_name='Indice_BET')

    workbook = writer.book
    worksheet = writer.sheets['Indice_BET']

    merge_format = workbook.add_format(
    {
        "bold": 1,
        "border": 1,
        "align": "center",
        "valign": "vcenter",
    }
    )

    worksheet.merge_range("A1:H2", "BET® (BUCHAREST EXCHANGE TRADING)", merge_format)
    worksheet.merge_range("I1:N2", "DETALII CONSTITUENTI / ZI", merge_format)
    number_format = workbook.add_format({'num_format': '#,##0.000'})

    for col_num in range(3, 13):
        worksheet.set_column(col_num, col_num, None, number_format)

    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 15)
    
workbook.close
