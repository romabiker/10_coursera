from bs4 import BeautifulSoup
from io import BytesIO
from lxml import etree
from openpyxl import Workbook
import argparse
import requests
import sys


def create_parser():
    parser = argparse.ArgumentParser()
    parser.add_argument('-k', '--keyword', type=str, required=True,
                        help='keyword for courses search')
    parser.add_argument('-c', '--count', default=5, type=int,
                        help='number of courses to search')
    parser.add_argument('-p', '--path', default='courses.xlsx', type=str,
                        help='file to save results')
    return parser


def find_commitment(soup):
    commitment_tr = soup.find('table', class_='basic-info-table').\
                    find('tbody').find(text='Commitment')
    if not commitment_tr:
        return 'not found'
    commitment_tag = commitment_tr.findNext('td')
    if not commitment_tag:
        return 'not found'
    return commitment_tag.text


def find_language(soup):
    language_tag = soup.find('div', class_='rc-Language')
    if not language_tag:
        return 'not found'
    return language_tag.text


def find_ratings(soup):
    ratings_tag = soup.find('div', class_='ratings-text')
    if not ratings_tag:
        return 'not found'
    return ratings_tag.text


def find_start_date(soup):
    start_date_tag = soup.find('div', class_='rc-StartDateString').find('span')
    if not start_date_tag:
        return 'not found'
    return start_date_tag.text


def find_title(soup):
    title_tag = soup.find('h1', class_='title display-3-text')
    if not title_tag:
        return 'not found'
    return title_tag.text


def get_coursera_xml():
    response = requests.get('https://www.coursera.org/sitemap~www~courses.xml')
    if not response.ok:
        response.raise_for_status()
    return response.text


def cook_courses_info(course_url):
    courses_info = []
    for url in courses_urls:
        course_html = request_course_html(url)
        soup = BeautifulSoup(course_html, 'html.parser')
        courses_info.append((
            find_title(soup),
            find_commitment(soup),
            find_language(soup),
            find_start_date(soup),
            find_ratings(soup),
            url,
        ))
    return courses_info


def output_courses_info_to_xlsx(courses_info, filepath):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append([
        'Course name',
        'Commitment',
        'Language',
        'Start date',
        'Course ratings',
        'Course url'
    ])
    for row in courses_info:
        worksheet.append(row)
    workbook.save(filepath)


def prepare_courses_urls(keyword, count):
    courses_xml_str = get_coursera_xml()
    courses_xml_io = BytesIO(bytes(courses_xml_str, encoding='utf-8'))
    return [element.text for event, element in etree.iterparse(courses_xml_io)
            if keyword in element.text][:count]


def request_course_html(course_url):
    response = requests.get(course_url)
    if not response.ok:
        response.raise_for_status()
    return response.content.decode('utf-8', 'ignore')


if __name__ == '__main__':
    parser = create_parser()
    namespace = parser.parse_args(sys.argv[1:])
    courses_urls = prepare_courses_urls(namespace.keyword, namespace.count)
    courses_info = cook_courses_info(courses_urls)
    output_courses_info_to_xlsx(courses_info, namespace.path)
