import requests
from bs4 import BeautifulSoup
import pandas as pd
import xlsxwriter

url ="https://bvb.ro/FinancialInstruments/Indices/IndicesProfiles.aspx?i=BET"

page=requests.get(url)

soup=BeautifulSoup(page.text,"html.parser")

table=soup.find('table',attrs={"class":"table table-hover dataTable no-footer generic-table compact"})

rows=table.find_all("tr")

headers = [header.text.strip() for header in rows[0].find_all('th')]

data = []
for row in rows[1:]:
    cells = row.find_all('td')
    data.append([cell.text.strip() for cell in cells])

df = pd.DataFrame(data, columns=headers)

df.to_excel('ConstituentiBet.xlsx', index=False)