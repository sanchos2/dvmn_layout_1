from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import os
import pandas


def excel_to_dict(excel_file, item_to_sort):
    excel_data_df = pandas.read_excel(excel_file, na_values=['nan'], keep_default_na=False)
    raw_wine_dict = excel_data_df.to_dict(orient='records')
    excel_dict = defaultdict(list)
    for item in raw_wine_dict:
        excel_dict[item[item_to_sort]].append(item)
    return excel_dict


path_to_excel = os.path.join('files', 'wine3.xlsx')
wine_dict = excel_to_dict(path_to_excel, 'Категория')

winery_start = datetime(year=1920, month=1, day=1)
date_now = datetime.now()
age_winery = date_now.year - winery_start.year

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')
rendered_page = template.render(
    age_winery=age_winery,
    wine_dict=wine_dict,
)

with open('index.html', 'w', encoding='utf8') as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8001), SimpleHTTPRequestHandler)
server.serve_forever()
