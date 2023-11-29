from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time
import re

# Укажите путь к вашему веб-драйверу
webdriver_path = '/path/to/chromedriver'

# Опции браузера (можно настроить)
chrome_options = Options()
chrome_options.add_argument('--headless')  # Запуск браузера в фоновом режиме
chrome_options.add_argument(f'--webdriver={webdriver_path}')  # Путь к драйверу

# Создаем объект браузера
driver = webdriver.Chrome(options=chrome_options)

# Замените URL на страницу с нужными котировками на TradingView
url = 'https://ru.tradingview.com/screener/'
driver.get(url)

# Подождем некоторое время для полной загрузки страницы
time.sleep(1)

# Получаем исходный код страницы после загрузки JavaScript
page_source = driver.page_source

# Закрываем браузер
driver.quit()

# Парсим HTML-код страницы
soup = BeautifulSoup(page_source, 'html.parser')

strings = soup.find_all('tr', class_='tv-data-table__row tv-data-table__stroke tv-screener-table__result-row')
namesCompanies = soup.find_all('span', class_='tv-screener__description')
prices = soup.find_all('td', class_='tv-data-table__cell tv-screener-table__cell tv-screener-table__cell--with-marker', attrs={'title': '', 'data-field-key': 'close'})
atrs = soup.find_all('td', class_='tv-data-table__cell tv-screener-table__cell tv-screener-table__cell--down tv-screener-table__cell--with-marker', attrs={'title': '', 'data-field-key': 'change'})

# Создаем список данных
data = []
# Заполнение списка данными
for name, price, atr in zip(namesCompanies, prices, atrs):
    data.append([name.text.strip(), price.text.strip(), atr.text.strip()])

# Создаем DataFrame с помощью Pandas
df = pd.DataFrame(data, columns=['Актив', 'Цена', 'ATR'])

# Сохраняем DataFrame в файл Excel
df.to_excel('котировки.xlsx', index=False)
