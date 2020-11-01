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
        required=True,
    )
    return parser


def convert_product_table(xlsx_file, key_to_sort):
    """
    Convert xslx file to dictionary.

    :param xlsx_file: xlsx file.
    :param key_to_sort: Key to sorting.
    :return: prepared dict.
    """
    product_table = pandas.read_excel(xlsx_file, na_values=['nan'], keep_default_na=False)
    raw_products = product_table.to_dict(orient='records')
    sorted_products = defaultdict(list)
    for record in raw_products:
        sorted_products[record[key_to_sort]].append(record)
    return sorted_products


winery_creation_year = 1920
winery_age = datetime.now().year - winery_creation_year

parser = create_parser()
namespace = parser.parse_args()
path_to_file = namespace.file
product_dict = convert_product_table(path_to_file, 'Категория')

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
