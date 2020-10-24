"""Main file."""
import os
from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def excel_to_dict(excel_file, item_to_sort):
    """
    Excel to dictionary function.

    :param excel_file: Excel file.
    :param item_to_sort: Key to sorting.
    :return: prepared dict.
    """
    excel_data_df = pandas.read_excel(excel_file, na_values=['nan'], keep_default_na=False)
    raw_wine_dict = excel_data_df.to_dict(orient='records')
    excel_dict = defaultdict(list)
    for wine in raw_wine_dict:
        excel_dict[wine[item_to_sort]].append(wine)
    return excel_dict


winery_creation_year = 1920
winery_creation_month = 1
winery_creation_day = 1
winery_start = datetime(
    year=winery_creation_year,
    month=winery_creation_month,
    day=winery_creation_day,
)
date_now = datetime.now()
age_winery = date_now.year - winery_start.year

path_to_excel = os.path.join('files', 'wine3.xlsx')
wine_dict = excel_to_dict(path_to_excel, 'Категория')

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml']),
)

template = env.get_template('template.html')
rendered_page = template.render(
    age_winery=age_winery,
    wine_dict=wine_dict,
)

with open('index.html', 'w', encoding='utf8') as main_html_file:
    main_html_file.write(rendered_page)

server = HTTPServer(('127.0.0.1', 8001), SimpleHTTPRequestHandler)
server.serve_forever()
