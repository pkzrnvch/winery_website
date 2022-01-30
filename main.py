from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import collections


def calculate_year():
    now = datetime.datetime.now()
    year_number = now.year - 1920
    if year_number < 105:
        year_text = "года"
    else:
        year_text = "лет"
    return year_number, year_text


def read_excel_file(filename):
    products_data = pandas.read_excel(filename)
    products_data = products_data.fillna('').to_dict(orient='records')
    products_data = sorted(products_data, key=lambda i: i['Категория'])
    structured_products_data = collections.defaultdict(list)
    for el in products_data:
        structured_products_data[el['Категория']].append(el)
    return structured_products_data


def render_page(year_info, structured_products_data):
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')
    rendered_page = template.render(
        year_number=year_info[0],
        year_text=year_info[1],
        structured_products_data=structured_products_data
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)


def main():
    render_page(calculate_year(), read_excel_file('wine3.xlsx'))

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
