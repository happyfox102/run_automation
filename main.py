import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time

# Читаем Excel
df = pd.read_excel('ваш_файл.xlsx')  # столбцы: Фамилия, Имя, Отчество, ДатаРождения и т.д.

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

for index, row in df.iterrows():
    driver.get('localhost')  # ссылка на форму

    # Примеры заполнения (замените селекторы под вашу форму!)
    driver.find_element(By.NAME, "fam").send_keys(row['Фамилия'])
    driver.find_element(By.NAME, "nam").send_keys(row['Имя'])
    driver.find_element(By.NAME, "otch").send_keys(row['Отчество'])
    driver.find_element(By.NAME, "birthdate").send_keys(row['ДатаРождения'])  # формат dd.mm.yyyy

    # Чекбокс "Розыск прекращен" если нужно
    # if row['РозыскПрекращен']: driver.find_element(By.ID, "checkbox_id").click()

    driver.find_element(By.XPATH, "//button[contains(text(), 'Поиск')]").click()

    time.sleep(3)  # подождать результаты, можно добавить скриншот или парсинг результатов

    print(f"Обработана строка {index + 1}")

driver.quit()