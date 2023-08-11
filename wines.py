import requests
from bs4 import BeautifulSoup
import pandas as pd


base_url = 'https://www.thewhiskyexchange.com/'
agent = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'}

wineLinks = []


def get_wineLinks():
    for x in range(1, 3):
        url = f'https://www.thewhiskyexchange.com/c/317/indian-whisky?pg={x}'
        r = requests.get(url, headers=agent)
        soup = BeautifulSoup(r.content, 'html.parser')
        wine_cards = soup.find_all('li', class_="product-grid__item")

        for items in wine_cards:
            for links in items.find_all('a', href=True):
                wineLinks.append(base_url+links['href'])

# print(wineLinks)


def wine_information():
    for links in wineLinks:
        r = requests.get(links, headers=agent)
        soup = BeautifulSoup(r.content, 'html.parser')

        divs = soup.find_all('article', class_='product-page')

        for items in divs:
            try:
                try:
                    wine_Names = items.find(
                        'h1', class_='product-main__name').text.strip().replace("\n", " ")
                except:
                    wine_Names = ""

                try:
                    wine_Description = items.find(
                        'div', class_='product-main__description').text.strip().replace("\n", " ")
                except:
                    wine_Description = ""
                try:
                    wine_Type = items.find(
                        'ul', class_='product-main__meta').text.strip().replace("\n", " ")
                except:
                    wine_Type = ""

                wineDicts = {
                    'Name': wine_Names,
                    'Type': wine_Type,
                    'Description ': wine_Description
                }
                list_of_Wines.append(wineDicts)
            except Exception as e:
                print(e)


list_of_Wines = []

get_wineLinks()
wine_information()
# print(list_of_Wines)
# print(len(list_of_Wines))
# print(list_of_Wines)

df = pd.DataFrame(list_of_Wines)
# print(df.head())
df.to_csv('wines.csv')
excel_writer = pd.ExcelWriter('wines.xlsx', engine='openpyxl')

# Convert the DataFrame to an XlsxWriter Excel object
df.to_excel(excel_writer, sheet_name='Sheet1', index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = excel_writer.book
worksheet = excel_writer.sheets['Sheet1']

# Iterate through all columns and set the width based on the maximum content length
for column in worksheet.columns:
    max_length = 0
    column = [cell for cell in column]
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

# Save the Excel file
excel_writer._save()
