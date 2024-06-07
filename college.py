import requests
from bs4 import BeautifulSoup
import openpyxl
import re
import os

def main():
    date = input("Enter date (MM/DD/YYYY): ")
    date_cleaned = re.sub(r'[^a-zA-Z0-9]', '_', date)  
    page = requests.get("http://www.isx-iq.net/isxportal/portal/homePage.html")
    view = page.content
    soup = BeautifulSoup(view, features="html.parser")
    stock_exchange_details = []

   
    movement_data = soup.find("div", {"id": "divDisplayTags", "class": "movement-data"})
    if movement_data:
        movement_rows = movement_data.find_all("div", {"class": "movement-row"})
        for row in movement_rows:
            company_name = row.find("div", {"class": "movement-cell1"}).text.strip()
            price = row.find("div", {"class": "movement-cell2"}).text.strip()
            change = row.find("div", {"class": "movement-cell4"}).text.strip()
            stock_exchange_details.append({"Company Name": company_name, "Price": price, "Change": change})

    
    market_data = soup.find("div", {"class": "summary-box"})
    if market_data:
        market_rows = market_data.find_all("div", {"class": "datarow1"})
        for row in market_rows:
            datacell1 = row.find("div", {"class": "datacell1"}).text.strip()
            datacell2 = row.find("div", {"class": "datacell2"}).text.strip()
            stock_exchange_details.append({"Market Data": datacell1, "Value": datacell2})

    
    save_to_excel(stock_exchange_details, date_cleaned)

def save_to_excel(stock_exchange_details, date):
    current_dir = os.getcwd()  
    file_path = os.path.join(current_dir, f"output_{date}.xlsx")  
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Data Type", "Value"])
    for item in stock_exchange_details:
        ws.append([item.get("Company Name", item.get("Market Data")), item.get("Price", item.get("Value")), item.get("Change", "")])
    wb.save(file_path)
    print(f'Data saved to {file_path}.')

main()






