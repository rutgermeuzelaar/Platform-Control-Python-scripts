import bs4 as bs4
import lxml as lxml
import requests
import pandas as pd
import openpyxl
import os

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:111.0) Gecko/20100101 Firefox/111.0'}
list_with_tags = ['Merk', 'EAN', 'Artikelcode', 'Gewicht', 'Hoogte', 'Lengte', 'Breedte', 'Maat', 'Inhoud', 'Omdoos',
                  'Verkoopprijs', 'Minimale bestelhoeveelheid', 'BTW']

dataframe_structure = {column: [] for column in list_with_tags}
dataframe_structure['product_url'] = []
print(dataframe_structure)
start_number = 15500
list_urls = []
status = []
dataframe_status_codes = {'status': status, 'url': list_urls}

for number in range(1, 50):
    url = f'https://www.hollandanimalcare.nl/product/{start_number}'
    productpage_request = requests.get(url, headers=headers)

    print('status code: ' + str(productpage_request.status_code))
    print(start_number)
    list_urls.append(url)
    status.append(str(productpage_request.status_code))
    start_number += 1

    if productpage_request.status_code == 200:
        dataframe_structure.get('product_url').append(url)
        for tag in list_with_tags:
            try:
                label_tag = bs4.BeautifulSoup(productpage_request.content, 'lxml').find('label', string=tag)
                next_tag = label_tag.find_next()
                label_tag_str = label_tag.text.strip()
                next_tag_str = next_tag.text.strip()
                dataframe_structure.get(label_tag_str).append(next_tag_str)

            except AttributeError:
                dataframe_structure.get(tag).append('')

# Export the dataframes to Excel
current_directory = os.path.dirname(__file__)
dataframe_status = pd.DataFrame(dataframe_status_codes)
dataframe_status.to_excel(os.path.join(current_directory, 'webscrape_product_info/export_status_codes.xlsx')
                          , index=False)

dataframe_export = pd.DataFrame(dataframe_structure)
dataframe_export.to_excel(os.path.join(current_directory, 'webscrape_product_info/export_product_information.xlsx')
                          , index=False)
