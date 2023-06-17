from selenium.webdriver import Chrome
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from time import sleep
from bs4 import BeautifulSoup
import bs4
from datetime import datetime
import requests
import xlsxwriter
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

MAIN_URL = 'https://moskva.gde.ru/'
servise = Service('chromedriver\chromedriver.exe')

class Parser():
    def __init__(self) -> None:
        options = webdriver.ChromeOptions()
        options.add_argument('--user-data-dir=C:\\Users\\danai\\AppData\\Local\\Google\\Chrome\\User Data\\Default')
        options.add_argument('profile-directory=Default')
        self.browser = Chrome(service=servise, options=options)
        self.browser.maximize_window()

    @staticmethod
    def get_all_main_categories()-> list:
        category_list = list()
        response = requests.get(url = MAIN_URL)
        html_page = BeautifulSoup(response.text, 'lxml')
        all_categories_html = html_page.find('ul', class_='nav-list', id='show-cat')
        for category_html in all_categories_html:
            if type(category_html) == bs4.element.Tag:
                category_url = category_html.find('a').get('href')
                category_list.append(category_url)
        return category_list
    
    def parsing_categories_page(self)-> list:
        category_list =  self.get_all_main_categories() #['https://krasnodar.gde.ru/nedvizhimost', ]
        all_items_url = list() 
        session = requests.Session()
        for category_url in category_list:
            for num_page in range(1, 2):
                full_category_url = category_url + f'?page={num_page}'
                print(full_category_url)
                response = session.get(full_category_url)
                html_page = BeautifulSoup(response.text, 'lxml')
                all_items_html = html_page.find('ul', class_='product-list').find_all('li')
                for item_html in all_items_html:
                    if item_html.find('script'):
                        continue
                    else:
                        item_url = item_html.find('div', class_='img-holder').find('a').get('href')
                        all_items_url.append(item_url)
        return all_items_url
    
    def parsing_all_items(self):
        all_items_info = []
        all_items_url = self.parsing_categories_page()
        for item_url in all_items_url:
            item_info = self.parsing_item_page(item_url=item_url)
            if item_info == 404:
                continue
            all_items_info.append(item_info)
        return all_items_info

    def parsing_item_page(self, item_url):
        print(item_url)
        self.browser.get(item_url)
        sleep(1)
        html_page = BeautifulSoup(self.browser.page_source, 'lxml')
        try:
            if html_page.find('div', id='main').find('h1').text == 'Ошибка 404':
                return 404
        except:
            pass
        item_name = html_page.find('h1', class_='product-name', itemprop='name').text 
        if html_page.find('div', class_='date') is not None:
            item_data = html_page.find('div', class_='date').text.replace('с ', '')
        else:
            item_data = 'отсутствует'
        if html_page.find('span', class_='address', itemprop='address') is not None:
            item_address = html_page.find('span', class_='address', itemprop='address').text
        else:
            item_address = 'отсутствует'

        item_categories = self.get_item_categories(html_page)
        item_price = self.get_item_price(html_page)
        item_options = self.get_item_options(html_page)

        item_description = self.get_item_description(html_page)
        try:
            item_owner_name = html_page.find('a', class_='user', itemprop='name').find('img').get('alt')
        except:
            item_owner_name = html_page.find('a', class_='user', itemprop='name').text.replace('\n', '')
        print(item_owner_name)

        item_owner_phone = self.get_item_owner_phone()
        item_img_list = self.get_item_img_list()

        return [item_name, item_price, item_data, item_address, item_description, item_categories,\
                item_options, item_img_list, item_owner_name, item_owner_phone]

    @staticmethod
    def get_item_categories(html_page):
        item_categoryes = str()
        for category_html in html_page.find_all('li', itemprop='itemListElement'):
            item_categoryes += category_html.find('span', itemprop='name').text + '|'
        return item_categoryes[0:-1]
    
    @staticmethod
    def get_item_price(html_page):
        if html_page.find('span', class_='price') is not None:
            item_price = html_page.find('span', class_='price').text
        else:
            item_price = 'отсутствует'
        return item_price
    
    @staticmethod
    def get_item_options(html_page):
        item_options = str() 
        if html_page.find('ul', class_='feature-list') is not None:
            for option_html in html_page.find('ul', class_='feature-list').find_all('dl'):
                option_name = option_html.find('dt').text
                option_value = option_html.find('dd').text
                item_options += f'{option_name} {option_value}|'
        else:
            item_options = 'отсутствует'
        return item_options
    
    @staticmethod
    def get_item_description(html_page):
        item_description = str()
        if html_page.find('div', class_='description', itemprop='description'):
            for description_html in html_page.find('div', class_='description', itemprop='description').find_all('p'):
                item_description += ' ' + description_html.text
        else:
            item_description = 'отсутствует'
        return item_description.replace('\xa0', '')
    
    def get_item_owner_phone(self):
        try:
            self.browser.find_element(By.XPATH, '//*[@id="sidebar"]/ul/li[1]/noindex/div[1]/div')
        except:
            return 'отсутствует'
        try:
            while True:
                try:
                    self.browser.find_element(By.XPATH, '//*[@id="sidebar"]/ul/li[1]/noindex/div[1]/div').click()
                    sleep(2)
                    html_page = BeautifulSoup(self.browser.page_source, 'lxml')
                    item_owner_phone = html_page.find('a', class_='tel').text
                    if 'Показать телефон' in item_owner_phone or 'Ошибка' in item_owner_phone or 'Загрузка' in item_owner_phone:
                        self.browser.refresh()
                        sleep(1)
                        print('обновил')
                        continue
                    print(item_owner_phone)
                    break
                except Exception as exc:
                    print(exc)
        except:
            pass
        return item_owner_phone
    
    def get_item_img_list(self):
        item_img_list = str()
        html_page = BeautifulSoup(self.browser.page_source, 'lxml')
        if html_page.find('div', id='bx-pager', class_='pager') is None:
            if html_page.find('table', class_='slide-img'):
                return html_page.find('table', class_='slide-img').find('img', itemprop='image').get('src')
            return 'отсутствует'
        max_num_img = len(html_page.find('div', id='bx-pager', class_='pager').find_all('div', class_='item'))
        try: 
            try:
                self.browser.find_element(By.XPATH, '//*[@id="main"]/div[4]/div/div[2]/div[1]/div[1]/div[1]/ul/li[2]/a[1]/table/tbody/tr/td/picture/img').click()
            except:
                self.browser.find_element(By.XPATH, '//*[@id="main"]/div[4]/div/div[2]/div[1]/div[1]/div[1]/ul/li[2]/a[1]/table/tbody/tr/td/img').click()
        except Exception as exc:
            print(exc)
        actions = ActionChains(self.browser)
        sleep(0.5)

        for _ in range(max_num_img):
            html_page = BeautifulSoup(self.browser.page_source, 'lxml')
            try:
                try:
                    img = html_page.find('div', class_='fancybox-slide fancybox-slide--html fancybox-slide--current fancybox-slide--complete')\
                                    .find('source', type='image/jpeg').get('srcset')
                except:
                    img = html_page.find('td', align='center').find('img').get('src')
            except Exception as exc:
                print(exc)
            item_img_list += img + '|'
            actions.send_keys(Keys.ARROW_RIGHT).perform()
            sleep(0.1)
        return item_img_list[0:-1]

    def write_items(self):
        all_item_info = self.parsing_all_items()

        book = xlsxwriter.Workbook(f'Объявления.xlsx')
        page = book.add_worksheet('объявления')

        row = 0
        column = 0

        page.set_column('A:A', 50)
        page.set_column('B:B', 50)
        page.set_column('C:C', 50)
        page.set_column('D:D', 50)
        page.set_column('E:E', 50)
        page.set_column('F:F', 50)
        page.set_column('G:G', 50)
        page.set_column('H:H', 50)
        page.set_column('I:I', 50)
        page.set_column('J:J', 50)

        page.write(row, column, 'название')
        page.write(row, column+1, 'цена')
        page.write(row, column+2, 'дата')
        page.write(row, column+3, 'адрес')
        page.write(row, column+4, 'описание')
        page.write(row, column+5, 'категории')
        page.write(row, column+6, 'опции')
        page.write(row, column+7, 'фото')
        page.write(row, column+8, 'владелец')
        page.write(row, column+9, 'телефон')

        row += 1

        for item_info in all_item_info:
            page.write(row, column, item_info[0])
            page.write(row, column+1, item_info[1])
            page.write(row, column+2, item_info[2])
            page.write(row, column+3, item_info[3])
            page.write(row, column+4, item_info[4])
            page.write(row, column+5, item_info[5])
            page.write(row, column+6, item_info[6])
            page.write(row, column+7, item_info[7])
            page.write(row, column+8, item_info[8])
            page.write(row, column+9, item_info[9])

            row += 1

        book.close()


if __name__ == '__main__':
    start = datetime.now()
    parser = Parser()
    parser.write_items()
    print(datetime.now() - start)
