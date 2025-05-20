from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import openpyxl
from urllib.request import urlopen
import webbrowser
from io import BytesIO

# Указываем путь до исполняемого файла Yandex Browser
yandex_browser_path = "/opt/yandex/browser/yandex-browser"

# Указываем путь до ChromeDriver
chromedriver_path = "/home/dmitriy/bin/yandexdriver"

# Настройки
options = Options()

#options.add_argument("--headless") # без запуска браузера
options.add_argument("user-agent=Mozilla/5.0 (X11; CrOS x86_64 10066.0.0) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36")
options.add_argument('--disable-cache')  # отключает кэш браузера

options.binary_location = yandex_browser_path

# Создаем драйвер
service = Service(executable_path=chromedriver_path)
browser = webdriver.Chrome(service=service, options=options)

browser.maximize_window()

# Зададим время на загрузку страницы, если по истечении этого времени страница не загрузилась, нажимаем ESC
start_download = time.perf_counter()
try:
    browser.set_page_load_timeout(30)
    browser.get("https://sletat.ru/hot-tour/")    
except:
    print("__________time out____________")
    browser.find_element(By.TAG_NAME, 'body').send_keys(Keys.CONTROL +'Escape')
# Конец загрузки
end_download = time.perf_counter()
print("Время потраченное на загрузку страницы = ", end_download - start_download, "секунд")

wb = openpyxl.Workbook()
list_cityes = []


start_time = time.perf_counter()  # время начала записи таблицы
while True:
    # Нажмем на меню Откуда
    browser.find_element(By.CSS_SELECTOR, 'div[class*="departure-city-select_departureCityContainer__vrrg5"]').click()
    time.sleep(3)
    print("Нажали на меню Откуда")

    #  Прокрутим выпадающее меню <ul class="sc-hCPjZK drrkbu sletatUIComponents2_select_optionsList"> до последнего пункта
    menu_cityes = browser.find_element(By.CSS_SELECTOR, 'ul[class*="sc-hCPjZK"]')
    end_city = browser.find_elements(By.CSS_SELECTOR, 'li[role="button"]')[-1]
    ActionChains(browser).click_and_hold(menu_cityes).move_to_element(end_city).release().perform()

    # Выбираем город в меню Откуда
    cityes = browser.find_elements(By.CSS_SELECTOR, 'li[role="button"]')
    if len(cityes) == len(list_cityes):
        break
    for city in cityes:
        if city.text not in list_cityes:
            print("Будем записывать = ", city.text)
            list_cityes.append(city.text)
            city.click()
            time.sleep(3)
            break
    
    country = browser.find_element(By.CSS_SELECTOR, 'div[class*="country-select_countryContainer__n3_Vq"] input[type="text"]').get_attribute('value')
    #print("country=", country)

    # В первый раз прокрутим страницу до самого низа (12 отелей), чтобы появилась кнопка Показать ещё
    browser.execute_script("window.scrollTo(0, document.body.scrollHeight);") 
    # Пауза, пока загрузится страница. 
    time.sleep(3)
    
    
    # ТАБЛИЦА hotel.xlsx
    # Создаём новую вкладку с именем = имени города
    if list_cityes[-1] == 'Москва':
        sheet = wb.active
        sheet.title = 'Москва'
    else:
        sheet = wb.create_sheet(title=list_cityes[-1])
    
    # Задаём ширину столбцов
    sheet.column_dimensions["A"].width = 17 # 17 символов
    sheet.column_dimensions["B"].width = 15
    sheet.column_dimensions["C"].width = 15
    sheet.column_dimensions["D"].width = 35  # непонятно в каких единицах, скорее всего в символах
    sheet.column_dimensions["E"].width = 15

    # Заголовок
    sheet.row_dimensions[1].height = 27  # высота ячейки в 27 пикселей
    sheet.append(["Город вылета", "Отель", "Страна тура", "Картинка отеля", "Минимальная цена", "Дата тура"])

    # Выравнивание текста по верху слева
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        for row in range(1, 60):
            sheet[f"{col}{row}"].alignment = openpyxl.styles.Alignment(horizontal="left", vertical="top", wrap_text=True)
        
      
    # Будем прокручивать страницу (нажимать Показать еще) пока количество отелей не будет > 36 или не исчезнет кнопка Показать
    count = 0
    while True:
        # Список отелей div class="tour-card_tourCardRoot__X7rNK finded-tours_gridTourCardItem__pGRj4">
        list_description_hotels = browser.find_elements(By.CSS_SELECTOR, 'div[class="tour-card_tourCardRoot__X7rNK finded-tours_gridTourCardItem__pGRj4"]')
            
        # Заполняем ниже заголовка
        for i in range(count, len(list_description_hotels)):
            print("i=", i+1, "count = ", count)
            row = i + 2
            sheet.row_dimensions[row].height = 200   # высота 200 пикселей

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

            # Получаем url изображения <img alt="" class="Slide_embla__slide__img__j26Zc"
            url = list_description_hotels[i].find_element(By.CSS_SELECTOR, 'img').get_attribute("src")
            print("url=", url)

            # Записываем jpg в xlsx без локального сохранения
            bufer = BytesIO(urlopen(url).read())
            if bufer.getvalue() == b'':
                sheet[f'D{row}'] = "Нет картинки"
            else:
                img = openpyxl.drawing.image.Image(bufer)
                img.height = 240
                img.width = 277
                sheet.add_image(img, f'D{row}')
    
            # Минимальная цена <span class="sc-hmdomO dffjpr">143 597</span>
            min_price = list_description_hotels[i].find_element(By.CSS_SELECTOR, 'span[class="sc-hmdomO dffjpr"]').text
            sheet[f'E{row}'] = min_price

            #Дата тура с <span class="Duration_date__MNy_I">27 апр</span> по <span class="Duration_date__MNy_I">3 мая</span
            data = list_description_hotels[i].find_elements(By.CSS_SELECTOR, 'span[class="Duration_date__MNy_I"]')
            sheet[f'F{row}'] = f'с {data[0].text} по {data[1].text}'
        
        # Если страницу прокрутили до количества > 36 отелей, то следующий город из меню Откуда иначе нажимаем Показать ещё
        count = len(list_description_hotels)
        print("Количество отелей =", count)   
        if count > 36:
            # Прокручиваем страницу вверх и начинаем новую итерацию из цикла Откуда
            browser.execute_script("window.scrollTo(document.body.scrollHeight, 0)")
            break
        else:
            try:
                # Нажимаем кнопку Показать еще
                WebDriverWait(browser, 25).until(EC.presence_of_element_located((By.XPATH, '//span[text()="Показать еще"]')))
                browser.find_element(By.XPATH, '//span[text()="Показать еще"]').click()
                time.sleep(4)
            except:
                print("Кнопки 'Показать еще' нет")
                break


# сохраняем все листы
wb.save("hotel.xlsx")
end_time = time.perf_counter()  # время конца записи таблицы
print("Время, потраченное на запись в hotel.xlsx = ", end_time - start_time, "секунд")

# Открываем hotel.xlsx
webbrowser.open('hotel.xlsx')
browser.quit()
