from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os
import openpyxl

filename = 'котировки.xlsx'
def checkRow(filename):
    df = pd.read_excel(filename, sheet_name='Sheet', header=None)
    last_row = df.iloc[:, 0].last_valid_index() + 1
    return last_row

def checkAndSaveFileExcel(filename, namesCompanies, prices, atrs):
    if os.path.isfile(filename):
        existing_data = pd.read_excel(filename, index_col=0)

        wb = openpyxl.load_workbook(filename)
        ws = wb.active
        last_row = checkRow(filename)

    else:
        wb = openpyxl.Workbook()
        ws = wb.active

        last_row = 0
        print('last_row', last_row)

    last_row = int(last_row)  # Преобразуем last_row в число

    for name, price, atr in zip(namesCompanies, prices, atrs):
        last_row += 1
        ws[f'A{last_row}'] = name.text.strip()
        ws[f'B{last_row}'] = price.text.strip()
        ws[f'C{last_row}'] = atr.text.strip()
    wb.save(filename)

def uploadInExcel(value):
    webdriver_path = '/path/to/chromedriver'
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Запуск браузера в фоновом режиме
    chrome_options.add_argument(f'--webdriver={webdriver_path}')  # Путь к драйверу
    driver = webdriver.Chrome(options=chrome_options)

    url = value
    driver.get(url)
    time.sleep(1)
    page_source = driver.page_source
    driver.quit()

    soup = BeautifulSoup(page_source, 'html.parser')

    namesCompanies = soup.find_all('span', class_='tv-screener__description')
    prices = soup.find_all('td',
                           class_='tv-data-table__cell tv-screener-table__cell tv-screener-table__cell--with-marker',
                           attrs={'title': '', 'data-field-key': 'close'})
    atrs = soup.find_all('td',
                         class_='tv-data-table__cell tv-screener-table__cell tv-screener-table__cell--down tv-screener-table__cell--with-marker',
                         attrs={'title': '', 'data-field-key': 'change'})

    data = []
    for name, price, atr in zip(namesCompanies, prices, atrs):
        data.append([name.text.strip(), price.text.strip(), atr.text.strip()])

    df = pd.DataFrame(data, columns=['Актив', 'Цена', 'ATR'])
    df.to_excel('котировки.xlsx', index=False)
    print('Успех!')

def uploadInExcelIndi(value):
    webdriver_path = '/path/to/chromedriver'
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Запуск браузера в фоновом режиме
    chrome_options.add_argument(f'--webdriver={webdriver_path}')  # Путь к драйверу
    driver = webdriver.Chrome(options=chrome_options)

    url = value
    driver.get(url)
    time.sleep(1)
    page_source = driver.page_source
    driver.quit()

    soup = BeautifulSoup(page_source, 'html.parser')

    namesCompanies = soup.find_all('h1', class_='apply-overflow-tooltip title-HFnhSVZy')
    prices = soup.find_all('span', class_='last-JWoJqCpY js-symbol-last')
    atrs = soup.find_all('span', class_='js-symbol-change-pt')

    checkAndSaveFileExcel(filename, namesCompanies, prices, atrs)
    print('Успех!')


# url = 'https://ru.tradingview.com/screener/'
urlSP500 = 'https://ru.tradingview.com/symbols/SPX/?exchange=SP'
urlHangSeng = 'https://ru.tradingview.com/symbols/TVC-HSI/'
urlIMOEX = 'https://ru.tradingview.com/symbols/MOEX-IMOEX/'
urlRTSI = 'https://ru.tradingview.com/symbols/MOEX-RTSI/'
urlGC1 = 'https://ru.tradingview.com/symbols/COMEX-GC1!/'
urlGOLDRUBTOM = 'https://ru.tradingview.com/symbols/USDRUB_TOM/?exchange=MOEX'
urlUcloilBrent = 'https://ru.tradingview.com/symbols/UKOIL/?exchange=TVC'
urlES1 = 'https://ru.tradingview.com/symbols/CME_MINI-ES11!/?contract=ES101Z2023'
urlNG1 = 'https://ru.tradingview.com/symbols/MOEX-NG1!/'
urlEURUSD = 'https://ru.tradingview.com/symbols/EURUSD/?exchange=OANDA'
urlCNHUSD = 'https://ru.tradingview.com/symbols/CNHUSD/?exchange=FX_IDC'
urlFGBL1 = 'https://ru.tradingview.com/symbols/EUREX-FGBL1!/'
urlRGBI1 = 'https://ru.tradingview.com/symbols/MOEX-RGBI/'
urlETHUSD = 'https://ru.tradingview.com/symbols/ETHUSD/?exchange=BINANCE'
urlBTCUSD = 'https://ru.tradingview.com/symbols/BTCUSD/?exchange=BINANCE'

uploadInExcelIndi(urlSP500)
uploadInExcelIndi(urlHangSeng)
uploadInExcelIndi(urlIMOEX)
uploadInExcelIndi(urlRTSI)
uploadInExcelIndi(urlGC1)
uploadInExcelIndi(urlGOLDRUBTOM)
uploadInExcelIndi(urlUcloilBrent)
uploadInExcelIndi(urlES1)
uploadInExcelIndi(urlNG1)
uploadInExcelIndi(urlEURUSD)
uploadInExcelIndi(urlCNHUSD)
uploadInExcelIndi(urlFGBL1)
uploadInExcelIndi(urlRGBI1)
uploadInExcelIndi(urlETHUSD)
uploadInExcelIndi(urlBTCUSD)