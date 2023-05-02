from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from bs4 import BeautifulSoup
import requests

web_page = requests.get('https://www.nts.live/infinite-mixtapes')
text = web_page.text

soup = BeautifulSoup(text, 'html.parser')

work_book = Workbook()
work_sheet = work_book.active

title_row = ['No.', 'Title', 'Description', 'Link']

work_sheet.append(title_row)

font = Font(bold=True)
alignment = Alignment(horizontal="left", vertical="center")

for row in work_sheet:
    for cell in row:
        cell.font = font
        cell.alignment = alignment

items = soup.find_all(class_='mixtape-tile-wrapper')

count = 0

for elem in items:
    count += 1
    title = elem.find(class_='mixtape-tile__detail-link__content__title__text').text
    description = elem.find(class_='mixtape-tile__detail-link__content__subtitle').text
    url = 'https://www.nts.live/infinite-mixtapes' + elem.find(class_='mixtape-tile__detail-link nts-app mobile-link').attrs['href']

    row = [count, title, description, url]
    print(row)

    work_sheet.append(row)

for row in work_sheet:
    for cell in row:
        cell.alignment = alignment

work_book.save('INFINITE MIXTAPES FROM NTS RADIO.xlsx')
