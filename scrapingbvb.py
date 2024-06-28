import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import threading
import os
from email.message import EmailMessage
import ssl
import smtplib
import json

global frecventa
frecventa={}
global initializator
initializator=0


def convert_to_float(value):
    try:
        clean_value = value.replace('.', '').replace(',', '.')
        return round(float(clean_value), 3)
    except ValueError:
        return value
    
def request_url(url):
    try:
        page = requests.get(url)
        return page
    except requests.exceptions.RequestException:
        print("Request exception occurred. Retrying...")
        time.sleep(1)
        return request_url(url)

def process_other(value):
    return value

def get_password():
    with open('config.json') as f:
        config = json.load(f)
        return config['password']

def trimite_mail(rand):
    email_sender="YOUR_MAIL"
    email_password=get_password()
    email_receiver="MAIL_RECEIVER"
    variatie=abs(rand[10])
    subject=[]
    if(variatie > 2 and variatie < 3):
        subject="Variatie usoara a "+rand[0]
    if(variatie > 3 and variatie < 5):
        subject="Variatie moderata a "+rand[0]
    if(variatie > 5):
        subject="Variatie extrema a "+rand[0]
    body = f""" 
    {rand[1]} a avut o variatie de pret de {rand[10]} % ajungand la pretul de {rand[8]}!
"""
    em=EmailMessage()
    em['From']=email_sender
    em['To']=email_receiver
    em['Subject']=subject
    em.set_content(body)

    context=ssl.create_default_context()

    with smtplib.SMTP_SSL('smtp.gmail.com',465,context=context) as smtp:
        smtp.login(email_sender,email_password)
        smtp.sendmail(email_sender,email_receiver,em.as_string())
    
def verifica_pret(rand):
    global frecventa
    if abs(rand[10])>2 and frecventa[rand[0]]==0:
        frecventa[rand[0]]+=1
        trimite_mail(rand)

def fetch_and_save_data(file_path, progress_bar):
    global initializator,frecventa
    url = "https://bvb.ro/FinancialInstruments/Indices/IndicesProfiles.aspx?i=BET"
    page=request_url(url)
    print("Connection established")
    soup = BeautifulSoup(page.text, "html.parser")

    table = soup.find('table', attrs={"class": "table table-hover dataTable no-footer generic-table compact"})
    rows = table.find_all("tr")
    pret_bet = soup.find("b", attrs={"class": "value"})
    pret_bet = pret_bet.text.strip()
    headers = [header.text.strip() for header in rows[0].find_all('th')]

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

    if initializator == 0:
        for i,row in enumerate(data):
            frecventa[data[i][0]]=0
        initializator+=1

    lung = len(data)
    ok = 0
    for i in range(lung):
        url = "https://bvb.ro/FinancialInstruments/Details/FinancialInstrumentsDetails.aspx?s=" + data[i][0]
        page = request_url(url)

        print("Connection established with "+data[i][0])
        soup = BeautifulSoup(page.text, "html.parser")
        table = soup.find('table', attrs={"id": "ctl00_body_ctl02_PricesControl_dvCPrices"})
        rows = table.find_all("tr")
        row_data = []
        for row in rows[5:12]:
            cells = row.find_all('td')
            for j, cell in enumerate(cells):
                cell_text = cell.text.strip()
                if ok == 0 and j % 2 == 0:
                    headers.append(cell_text)
                else:
                    if j % 2 != 0:
                        row_data.append(convert_to_float(cell_text))

        for row in rows[14:16]:
            cells = row.find_all('td')
            for j, cell in enumerate(cells):
                cell_text = cell.text.strip()
                if ok == 0 and j % 2 == 0:
                    headers.append(cell_text)
                else:
                    if j % 2 != 0:
                        row_data.append(convert_to_float(cell_text))
        if ok == 0:
            headers.append("Ultimul Dividend")
            headers.append("Grafic")
        table = soup.find('table', attrs={"id": "ctl00_body_ctl02_IndicatorsControl_dvIndicators"})
        rows = table.find_all("tr")
        nr = 0
        for row in rows:
            if nr == 0:
                cells = row.find_all("td")
                j = 0
                for k, cell in enumerate(cells):
                    cell_text = cell.text.strip()
                    if "Dividend" in cell_text:
                        j += 1
                    else:
                        if j > 0:
                            nr = nr + 1
                            row_data.append(convert_to_float(cell_text))
        if nr == 0:
            row_data.append(0)
        row_data.append("https://www.tradingview.com/chart/hHTcjp5L/?symbol=BVB%3A" + data[i][0])
        data[i].extend(row_data)
        ok += 1
        verifica_pret(data[i])
        progress_bar['value'] = (i+1) * 100 / lung
        root.update_idletasks()

    df = pd.DataFrame(data, columns=headers)

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, startrow=2, sheet_name='Indice_BET')

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
        font_red = workbook.add_format({'font_color': 'red', 'num_format': '#,##0.000'})
        font_green = workbook.add_format({'font_color': 'green', 'num_format': '#,##0.000'})

        col = 10
        for row in range(0, lung):
            cell_value = df.iloc[row, col]
            if isinstance(cell_value, (int, float)):
                if cell_value < 0:
                    worksheet.write(row + 3, col, cell_value, font_red)
                elif cell_value >= 0:
                    worksheet.write(row + 3, col, cell_value, font_green)
            else:
                worksheet.write(row + 3, col, cell_value)

        for col_num in range(3, 18):
            worksheet.set_column(col_num, col_num, 10, number_format)
        bold_format = workbook.add_format({'bold': True})

        worksheet.set_column(1, 1, 30)
        worksheet.set_column(2, 2, 15)
        worksheet.set_column(0, 0, 10, bold_format)
        pretul = convert_to_float(pret_bet)
        worksheet.write_number(lung + 4, 2, pretul)
        worksheet.merge_range(lung + 4, 0, lung + 4, 1, "PRET BET", merge_format)
    
    messagebox.showinfo("Succes", f"Fișierul a fost salvat la {file_path}")

def schedule_next_run():
    root.after(60000, start_fetch_and_save)

def start_fetch_and_save():
    global saved_file_path
    if saved_file_path is None:
        saved_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])
    else:
        print("//////////////Aplicatia s-a mai executat o data///////////////")
    if saved_file_path:
        progress_bar['value'] = 0
        thread = threading.Thread(target=fetch_and_save_data, args=(saved_file_path, progress_bar))
        thread.start()
        schedule_next_run()

def center_window(window, width=400, height=200):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()
    x = (screen_width // 2) - (width // 2)
    y = (screen_height // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

saved_file_path = None

root = tk.Tk()
root.title("Salvare Fișier BET")
root.configure(background="lightblue")

center_window(root, 400, 200)

button_select_path = tk.Button(root, text="Selectează calea și salvează fișierul", command=start_fetch_and_save)
button_select_path.pack(pady=20)

progress_bar = ttk.Progressbar(root, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=20)

root.mainloop()
