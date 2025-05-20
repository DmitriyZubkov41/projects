from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from time import time, sleep
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.drawing.image import Image 
import webbrowser
from io import BytesIO
from aiohttp import ClientSession
import asyncio
import logging

def open_page():
    
    # Указываем путь до исполняемого файла Yandex Browser
    yandex_browser_path = "/opt/yandex/browser/yandex-browser"
    # Указываем путь до ChromeDriver
    chromedriver_path = "/home/dmitriy/bin/yandexdriver"
    # Настройки
    options = Options()
    #options.add_argument("--headless") # без запуска браузера
    options.add_argument(
     "user-agent=Mozilla/5.0 (X11; CrOS x86_64 10066.0.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36"
    )
    options.add_argument('--disable-cache')  # отключает кэш браузера
    options.binary_location = yandex_browser_path
    # Создаем драйвер
    service = Service(executable_path=chromedriver_path)
    browser = webdriver.Chrome(service=service, options=options)
    browser.maximize_window()
    # Открываем страницу
    try:
        browser.set_page_load_timeout(30)
        browser.get("https://sletat.ru/hot-tour/")    
    except:
        print("__________time out____________")
        browser.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL +'Escape')
    
    return browser 


def get_city_in_menu(browser, list_cityes):
    page = browser.find_element(By.TAG_NAME, 'html')
    print("До выбора города, page.id=", page.id)
        
    # Нажмем на меню Откуда
    browser.find_element(By.CSS_SELECTOR, 'div[class*="departure-city-select_departureCityContainer__vrrg5"]').click()
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'ul[class^="sc-hCPjZK"]')))
    #  Прокрутим выпадающее меню ul class="sc-hCPjZK drrkbu sletatUIComponents2_select_optionsList" до последнего пункта
    menu_cityes = browser.find_element(By.CSS_SELECTOR, 'ul[class^="sc-hCPjZK"]')
    end_city = menu_cityes.find_elements(By.CSS_SELECTOR, 'li[role="button"]')[-1]
    ActionChains(browser).click_and_hold(menu_cityes).move_to_element(end_city).release().perform()

    # Выбираем город в меню Откуда
    cityes = browser.find_elements(By.CSS_SELECTOR, 'li[role="button"]')
    for city in cityes:
        if city.text not in list_cityes:
            list_cityes.append(city.text)
            print(f"Будем нажимать на  {city.text}")
            city.click()
            sleep(3)
            break

    return len(cityes)


def scroll_page(count_city, browser):
    while True:
        page = browser.find_element(By.TAG_NAME, 'html')
        print("page.id=", page.id)
        print("URL сессии =", browser.current_url)
        
        # Дождемся когда появится <h2 class="MuiTypography-root MuiTypography-h6">Популярные курорты:</h2>
        WebDriverWait(browser, 20).until(EC.presence_of_element_located((By.XPATH, '//h2[text()="Популярные курорты:"]')))
        # Прокручиваем в самый низ
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        # Проверим наличие кнопки Показать еще
        try:
            WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, '//span[text()="Показать еще"]')))
            
        except:
            # Нет кнопки, тогда:
            try:
                # 1. Если нет туров: <p class="MuiTypography-root jss9 MuiTypography-body1">Нет туров по ука
                no_tur = browser.find_element(By.XPATH, '//p[text()="Нет туров по указаным параметрам"]')
                print("Нет туров")
                return []
            except:
                try:
                    # 2. Нет элемента <span class="jss2">Ищем больше туров...</span> тогда выходим иначе снова прокручиваем
                    browser.find_element(By.XPATH, '//span[text()="Ищем больше туров..."]')
                    print("Страница не до конца загрузилась, снова прокрутим её")
                    continue
                except:
                    list_description_hotels = browser.find_elements(
                       By.CSS_SELECTOR, 'div[class="tour-card_tourCardRoot__X7rNK finded-tours_gridTourCardItem__pGRj4"]'
                     )
                    print("Страница загрузилась, но кнопки 'Показать еще' нет")
                    return list_description_hotels

        # Нажимаем на кнопку
        # Сначало прокрутим страницу к кнопке
        button = browser.find_element(By.XPATH, '//span[text()="Показать еще"]')
        ActionChains(browser).scroll_to_element(button).perform()
            
        # Будем нажимать Показать еще пока количество отелей не будет > count_city
        # или не исчезнет кнопка Показать
    
        # Список отелей div class="tour-card_tourCardRoot__X7rNK finded-tours_gridTourCardItem__pGRj4">
        list_description_hotels = browser.find_elements(
            By.CSS_SELECTOR, 'div[class="tour-card_tourCardRoot__X7rNK finded-tours_gridTourCardItem__pGRj4"]'
         )
        # Если страницу прокрутили до количества > 36 отелей, то записываем в таблицу иначе нажимаем Показать ещё
        count = len(list_description_hotels)
        if count > count_city:
            print("Количество отелей для записи =", count)
            return list_description_hotels
        else:
            # Нажимаем кнопку Показать еще
            print("Кнопка есть, но количество отелей меньше count_city, нажимаем на кнопку и прокручиваем страницу")
            button.click()
            continue
           

async def get_img_in_bufer(session, url):
    async with session.get(url) as response:
        image_in_bufer = BytesIO(await response.read())
        return image_in_bufer
            

async def main():
    start_time = time()  # время начала работы программы
    logging.basicConfig(level=logging.DEBUG, filename="sletat.log", filemode="w", format = "%(asctime)s - %(levelname)s - %(message)s")
    wb = Workbook()
    list_cityes = []
    count_city_in_menu = 0
    
    browser = open_page()

    while True:
        # Условие выхода из цикла
        if count_city_in_menu == len(list_cityes) and count_city_in_menu != 0:
            break

        # нажимаем на город в меню
        count_city_in_menu = get_city_in_menu(browser, list_cityes)
        
        # Прокручиваем страницу и жмём кнопку     
        list_description_hotels = scroll_page(36, browser)
        
        # Заполняем ТАБЛИЦУ hotel.xlsx
        urls = []
        # Создаём новую вкладку с именем = имени города
        if list_cityes[-1] == 'Москва':
            sheet = wb.active
            sheet.title = 'Москва'
        else:
            sheet = wb.create_sheet(title=list_cityes[-1])
    
        # Задаём ширину столбцов
        sheet.column_dimensions["A"].width = 17 
        sheet.column_dimensions["B"].width = 15
        sheet.column_dimensions["C"].width = 15
        sheet.column_dimensions["D"].width = 45
        sheet.column_dimensions["E"].width = 15
        
        # Заголовок
        sheet.row_dimensions[1].height = 30
        for col in [('A', 'Город вылета'), ('B', 'Отель'), ('C', 'Страна тура'), ('D', 'Картинка отеля'), 
                    ('E', 'Минимальная цена'), ('F', 'Дата тура')]:
            sheet[f"{col[0]}1"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            sheet[f"{col[0]}1"] = col[1]
        
        # Заполняем таблицу ниже заголовка кроме рисунков
        country = browser.find_element(
                    By.CSS_SELECTOR, 'div[class*="country-select_countryContainer__n3_Vq"] input[type="text"]'
        ).get_attribute('value')
        for i in range(len(list_description_hotels)):
            print("row=", i+2)
            row = i + 2
            sheet.row_dimensions[row].height = 200

            # Выравнивание текста по верху слева
            for col in ['A', 'B', 'C', 'E', 'F']:
                sheet[f"{col}{row}"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
            
            # Город вылета
            sheet[f'A{row}'] = list_cityes[-1]

            # Имя отеля
            name_hotel = list_description_hotels[i].find_element(By.CSS_SELECTOR, 'h3[class="HotelName_title__1D_6z"]')
            name_hotel = name_hotel.text.split('\n')
            if len(name_hotel) > 1:
                name_hotel = name_hotel[1]
            else:
                name_hotel = name_hotel[0]
            sheet[f'B{row}'] = name_hotel

            # Страна тура
            sheet[f'C{row}'] = country

            # Минимальная цена span class="sc-hmdomO dffjpr"
            min_price = list_description_hotels[i].find_element(By.CSS_SELECTOR, 'span[class="sc-hmdomO dffjpr"]').text
            sheet[f'E{row}'] = min_price

            #Дата тура с span class="Duration_date__MNy_I">27 апр по span class="Duration_date__MNy_I">
            data = list_description_hotels[i].find_elements(By.CSS_SELECTOR, 'span[class="Duration_date__MNy_I"]')
            sheet[f'F{row}'] = f'с {data[0].text} по {data[1].text}'
        
        # Заполняем таблицу картинками    
        urls = [hotel.find_element(By.CSS_SELECTOR, 'img').get_attribute("src") for hotel in list_description_hotels]
        async with ClientSession() as session:
            tasks = [get_img_in_bufer(session, url) for url in urls]
            images_bufers = await asyncio.gather(*tasks)

            for i in range(len(images_bufers)):
                if images_bufers[i].getvalue() == b'':
                    sheet[f'D{i+2}'] = "Нет картинки"
                else:
                    img = Image(images_bufers[i])
                    sheet.add_image(img, f'D{i+2}')
        
        
    browser.quit()
    # сохраняем все листы
    wb.save("hotel.xlsx")
    print("Время, потраченное на запись в hotel.xlsx = ", round((time() - start_time), 0), "seconds")
    # Открываем hotel.xlsx
    webbrowser.open('hotel.xlsx')
    

asyncio.run(main())
