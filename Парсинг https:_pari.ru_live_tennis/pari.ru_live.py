from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time
from yaml import safe_load

# Путь к Yandex Browser
YANDEX_BROWSER_PATH = "/opt/yandex/browser/yandex-browser"

# Настройки
options = Options()
options.binary_location = YANDEX_BROWSER_PATH

# WebDriver Manager сам подберет и установит нужный ChromeDriver
service = Service(ChromeDriverManager().install())

# Создаем WebDriver
browser = webdriver.Chrome(service=service, options=options)

browser.maximize_window()
browser.get("https://pari.ru/live/tennis")

# По какому счёту будем фильтровать, берётся из config.yml
with open('config.yml', 'r') as f:
  data = safe_load(f)
list_scores = data['scores'].split()

# Ждём когда загрузится круг чата
#div class="chat-button--kT6fd"><div class="wrap--iDLld"><span class="text--AIBKE">Задать вопрос</span>
WebDriverWait(browser, 30).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'div[class="chat-button--kT6fd"]')))

# Элемент панели, на который можно кликнуть без открытия меню, чтобы сфокусировать мышь на панель:
browser.find_elements(By.CSS_SELECTOR, 'span[class="event-block-score__score--r0ZU9"]')[1].click()

games = []
links = []
scores = []
win_1player = []
win_2player = []
length_games = 0

while True:
    # Линии: <div style="top: 52px; z-index: 22; position: sticky;">
    lines = browser.find_elements(By.CSS_SELECTOR, 'div[style^="top:"]')
    for line in lines:
        try:
            game = line.find_element(By.CSS_SELECTOR, 'div[class="sport-base-event__main--FHhdx"] a')
            if game.text not in games and line.find_element(By.CSS_SELECTOR, 'span[class="event-block-score__score--r0ZU9"]').text in list_scores:
                # Заносим в список games названия матчей
                games.append(game.text)
                # Список ссылок на матчи links
                links.append(game.get_attribute('href'))
                # Список на счёт в матче
                scores.append(line.find_element(By.CSS_SELECTOR, 'span[class="event-block-score__score--r0ZU9"]').text)
                # Список ставок на победу первого игрока
                after_score = line.find_elements(By.CSS_SELECTOR, 'div[class^="factor-value--zrkpK"]')
                win_1player.append(after_score[0].text)
                # Список ставок на победу второго игрока
                win_2player.append(after_score[2].text)
        except:
            continue
    # Условие, что прокрутили до конца, чтобы выйти из цикла:
    if length_games != len(games):
        length_games = len(games)
        
        # Нажмем на кнопку Page_Down:
        browser.find_element(By.TAG_NAME, 'html').send_keys(Keys.PAGE_DOWN)
        time.sleep(1)
        continue
    else: break
    
#print("Записано игр:", len(games))

# Запись в файл
text_file = open('tennis.txt', 'w')
text_file.write("Название матча, Счёт в матче, На 1 игрока-2 игрока, Ссылка на матч\n")

for i in range(len(games)):
    stroka = f'{games[i]}, {scores[i]}, {win_1player[i]}-{win_2player[i]}, {links[i]}\n'
    text_file.write(stroka)

text_file.close()
