import argparse
from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas
from jinja2 import Environment, FileSystemLoader, select_autoescape


def create_parser():
    """
    Parse command line option.

    :return: parser object
    """
    parser = argparse.ArgumentParser()
    parser.add_argument(
        '-f',
        '--file',
        help='Path to file with products. Format - xlsx.',
        type=str,
    )
    return parser


def xlsx_to_dict_converter(xlsx_file, key_to_sort):
    """
    Convert xslx file to dictionary.

    :param xlsx_file: xlsx file.
    :param key_to_sort: Key to sorting.
    :return: prepared dict.
    """
    excel_data_df = pandas.read_excel(xlsx_file, na_values=['nan'], keep_default_na=False)
    raw_dict = excel_data_df.to_dict(orient='records')
    sorted_dict = defaultdict(list)
    for record in raw_dict:
        sorted_dict[record[key_to_sort]].append(record)
    return sorted_dict


winery_creation_year = 1920
winery_creation_month = 1
winery_creation_day = 1
winery_start = datetime(
    year=winery_creation_year,
    month=winery_creation_month,
    day=winery_creation_day,
)
date_now = datetime.now()
winery_age = date_now.year - winery_start.year

parser = create_parser()
namespace = parser.parse_args()
path_to_xlsx = namespace.file
product_dict = xlsx_to_dict_converter(path_to_xlsx, 'Категория')

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml']),
)

template = env.get_template('template.html')
rendered_page = template.render(
    winery_age=winery_age,
    product_dict=product_dict,
)

with open('index.html', 'w', encoding='utf8') as main_html_file:
    main_html_file.write(rendered_page)

server = HTTPServer(('127.0.0.1', 8001), SimpleHTTPRequestHandler)
server.serve_forever()
