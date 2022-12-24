#!/usr/bin/env python
# coding: utf-8

import time
import re
import requests as req
import xlsxwriter as xl
from bs4 import BeautifulSoup
from selenium import webdriver
from datetime import datetime, timedelta


DRIVER = webdriver.Firefox(r"path_to_driver")
FILE = 'file.xlsx'
URL1 = "https://skidkaonline.ru/barnaul/{date}/shops/{shop}/?cids=469,630,2754,551"
URL2 = 'https://lenta.com/promo/'
URL3 = 'https://magnit.ru/promo/?category[]=krasota&category[]=byt_chim' 
URL4 = 'https://5ka.ru/special_offers/' 
HEADERS = {'user-agent':'user-agent'}


def write_excel(shop, content):
    """Создаёт новый лист в файле с именем магазина и
    вносит новые записи"""
    
    worksheet = workbook.add_worksheet(f'{shop.title()}')
    worksheet.write(0, 0, '№')
    worksheet.write(0, 1, 'Дата')
    worksheet.write(0, 2, 'Номенклатура')
    worksheet.write(0, 3, 'Цена со скидкой')
    worksheet.write(0, 4, '% Скидки')
    worksheet.write(0, 5, 'Цена без скидки')
    worksheet.write(0, 6, 'Период акции')
    
    row = 1
    for record in content:
        for elem in record:
            worksheet.write(row, record.index(elem), elem)
        row += 1
    
    print("Все записи внесены в файл.")


def collect_data_shop(driver, shop, url):
    dates = [(datetime.today()-timedelta(days=7*i)).strftime("%d-%m-%Y") for i in range(2)]
    shop_content = []
    i = 0
    
    for date in dates:
        driver.get(url.format(shop=shop, date=date))
        time.sleep(10) #Время для загрузки всех элементов
        items = driver.find_elements_by_css_selector("div.item")
        
        for item in items:
            i += 1
            try:
                name = item.find_element_by_tag_name("h3").text
                price = item.find_element_by_class_name("product-priceafter").text.replace(" руб.", "").replace(".", ",")
                discount = item.find_element_by_class_name("product-discount").text.replace("%", "").replace(".", ",")
                last_price = item.find_element_by_class_name("product-pricebefore-val").text.replace(".", ",")
                period = item.find_element_by_class_name("discount-link").text.replace('Акционные предложения\n(', "").replace(")", "")
                good = [shop, i, date, name, int(price), int(discount)/100, int(last_price), period]
            except:
                good += [shop, i, date, *["Н\Д" for i in range(5)]]
                
            shop_content += [good]
        
    return shop_content


def collect_data_lenta(driver, shop, url):
    driver.get(url)
    time.sleep(5)
    shop_content = []
    i = 0
    
    #Находим ссылки на категории по их названию
    categories_l = ['Красота и здоровье', 'Бытовая химия', 'Все для дома']
    categories_d = {}
    for category in categories_l:
        categories_d[category] = driver.find_element_by_link_text(category).get_attribute('href')

    for link in categories_d.values():
        driver.get(link)

        #Находим последнюю страницу
        pages = driver.find_elements_by_class_name('pagination__item')
        pages = [page.text for page in pages]
        last_page = int(pages[-1])
        print(f"Всего найдено {last_page} страниц.")

        for page in range(1, last_page+1):
            print(f"Парсим {page} страницу.")
            link = link + f'&page={page}'
            driver.get(link)

            items = driver.find_elements_by_class_name('sku-card-small-container')
            for item in items:
                i += 1
                name = item.find_element_by_class_name("sku-card-small__title").text
                prices = item.find_elements_by_class_name("sku-price__integer")
                price = prices[1].text
                last_price = prices[0].text
                discount = item.find_element_by_class_name('sku-card-small__labels').text
                discount = int(discount.replace("%", ''))/100

                #Находим период акции с помощью bs4
                period_link = item.find_element_by_tag_name("a").get_attribute("href")
                html = req.get(period_link, headers=HEADERS)
                soup = BeautifulSoup(html.text, 'html.parser')
                period_raw_text = soup.find('div', class_='sku-page__stocks-date').text
                periods = re.findall(r"\d\d\.\d\d\.\d\d\d\d", period_raw_text)
                period = periods[0]+'-'+periods[1]

                good = [shop, i, datetime.today().strftime("%d-%m-%Y"), name, price, discount, last_price, period]
                shop_content += [good]

    return shop_content


def collect_data_magnit(driver, shop, url):
    driver.get(url)
    time.sleep(5)
    shop_content = []
    i = 0
    
    #Подтверждаем, что нам есть 18 лет :)
    age_question = driver.find_element_by_class_name('confirm_age__answer')
    yes = age_question.find_elements_by_tag_name('button')
    yes[1].click()
    time.sleep(3)

    #Выбираем нужный нам город
    driver.find_element_by_class_name('header__contacts-link_city').click()
    search = driver.find_element_by_name('citySearch')
    search.clear()
    search.send_keys("г. Барнаул, Алтайский край" + Keys.ENTER)
    time.sleep(3) #Даём время сайту найти подходящий элемент
    driver.find_element_by_class_name('city-search__link').click()

    #Начинаем парсить
    items = driver.find_elements_by_class_name('card-sale_catalogue')

    for item in items:
        i += 1
        try:
            name = item.find_element_by_class_name("card-sale__title").text
            price = item.find_element_by_class_name("label__price_new").text
            price = price[:-2].replace('\n', '')+','+price[-2:]
            last_price = item.find_element_by_class_name("label__price_old").text
            last_price = last_price[:-2].replace('\n', '')+','+last_price[-2:]
            discount = item.find_element_by_class_name('card-sale__discount').text
            discount = int(discount.replace("%", '').replace('−', '-'))/100
            period = item.find_element_by_class_name('card-sale__date').text.replace('\n', '-').replace('с ', '').replace('до ', '')
            period = period +f' {datetime.today().strftime("%Y")} г.'
        except:
            name, price, last_price, discount, period = ["Н\Д" for i in range(5)]

        good = [shop, i, datetime.today().strftime("%d-%m-%Y"), name, price, discount, last_price, period]
        shop_content += [good]

    return shop_content


def collect_data_pyatorochka(driver, shop, url):
    driver.get(url)
    time.sleep(5)
    shop_content = []
    i = 0
    
    #Выбираем нужный нам город
    driver.find_element_by_class_name('location').click()
    search = driver.find_element_by_class_name('search__input')
    search.send_keys("г. Москва" + Keys.ENTER)
    time.sleep(5)
    driver.find_element_by_class_name('resultLine').click()

    #Подтверждаем, что согласны на куки
    button = driver.find_element_by_class_name("message__button")
    button.click()

    #Разворачиваем страницу
    time.sleep(3)
    button = driver.find_element_by_class_name("special-offers__more-btn")
    button.click()
    print('Разворачиваем страницу.')
    while button:
        try:
            time.sleep(3)
            button = driver.find_element_by_class_name("special-offers__more-btn")
            button.click()
        except:
            break

    time.sleep(5)
    items = driver.find_elements_by_class_name("sale-card")

    for item in items:
        i += 1
        name = item.find_element_by_class_name("sale-card__title").text
        price = item.find_element_by_class_name("sale-card__price--new").text
        price_raw = price[:int(len(price)/2)]
        price = price_raw[:-2]+','+price_raw[-2:]
        last_price_raw = item.find_element_by_class_name("sale-card__price--old").text
        last_price = last_price_raw[:-2]+','+last_price_raw[-2:]
        discount = round(float(price_raw[:-2]+'.'+price_raw[-2:])/float(last_price_raw[:-2]+'.'+last_price_raw[-2:])-1, 2)
        period = item.find_element_by_class_name('sale-card__date').text

        good = [shop, i, datetime.today().strftime("%d-%m-%Y"), name, price, discount, last_price, period]
        shop_content += [good]
    
    return shop_content


def main(driver, file):
    #Открываем файл Excel для записи
    workbook = xl.Workbook(file)
    content = []
    
    shops = ['ashan-2', 'yarche']
    for shop in shops:
        content += collect_data_shop(driver, shop, URL1)
        print(f'Парсинг {shop.title()} закончен.')
        
    shop = 'Лента'
    content += collect_data_lenta(driver, shop, URL2)
    print(f'Парсинг {shop.title()} закончен.')
    
    shop = 'Магнит'
    content += collect_data_magnit(driver, shop, URL3)
    print(f'Парсинг {shop.title()} закончен.')
        
    shop = 'Пятёрочка'
    content += collect_data_magnit(driver, shop, URL4)
    print(f'Парсинг {shop.title()} закончен.')
    
    #Сохраняем и закрываем Excel файл
    write_excel(content)
    workbook.close()
    print('Можно проверять файл.')

if __name__ == '__main__':
    main(DRIVER, FILE)

