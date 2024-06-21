import requests
from bs4 import BeautifulSoup
import pandas as pd

url = "https://bvb.ro/FinancialInstruments/Indices/IndicesProfiles.aspx?i=BET"

page = requests.get(url)
soup = BeautifulSoup(page.text, "html.parser")


table = soup.find('table', attrs={"class": "table table-hover dataTable no-footer generic-table compact"})
rows = table.find_all("tr")
pret_bet=soup.find("b",attrs={"class":"value"})
pret_bet=pret_bet.text.strip()
headers = [header.text.strip() for header in rows[0].find_all('th')]

def convert_to_float(value):
    try:
        clean_value=value.replace('.','').replace(',', '.')
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
    for row in rows[5:12]:
        cells = row.find_all('td')
        for j, cell in enumerate(cells):
            cell_text = cell.text.strip()
            if ok == 0 and j%2 == 0:
                headers.append(cell_text)
            else:
                if j%2 != 0:
                    row_data.append(convert_to_float(cell_text))

    for row in rows[14:16]:
        cells = row.find_all('td')
        for j, cell in enumerate(cells):
            cell_text = cell.text.strip()
            if ok == 0 and j%2 == 0:
                headers.append(cell_text)
            else:
                if j%2 != 0:
                    row_data.append(convert_to_float(cell_text))
    if ok == 0:
        headers.append("Grafic")
    row_data.append("https://www.tradingview.com/chart/hHTcjp5L/?symbol=BVB%3A"+data[i][0])
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
    worksheet.merge_range("I1:O2", "DETALII CONSTITUENTI / ZI", merge_format)
    worksheet.merge_range("P1:Q2", "MAXIME", merge_format)
    number_format = workbook.add_format({'num_format': '#,##0.000'})
    font_red = workbook.add_format({'font_color': 'red','num_format': '#,##0.000'})
    font_green = workbook.add_format({'font_color': 'green','num_format': '#,##0.000'})

    col = 10  # Indexul coloanei K
    for row in range(0, lung):
        cell_value = df.iloc[row, col]
        if isinstance(cell_value, (int, float)):
            if cell_value < 0:
                worksheet.write(row+3, col, cell_value, font_red)
            elif cell_value >= 0:
                worksheet.write(row+3, col, cell_value, font_green)
        else:
            worksheet.write(row+3, col, cell_value)

    for col_num in range(3, 17):
        worksheet.set_column(col_num, col_num, 10, number_format)
    bold_format = workbook.add_format({'bold': True})

    worksheet.set_column(1, 1, 30)
    worksheet.set_column(2, 2, 15)
    worksheet.set_column(0, 0, 10,bold_format)
    pretul=convert_to_float(pret_bet)
    worksheet.write_number(lung+4,2,pretul)
    worksheet.merge_range(lung+4,0,lung+4,1,"PRET BET",merge_format)