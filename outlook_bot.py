import os
from time import sleep
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver.driver import driver
from constants import credentials

driver.get('https://login.live.com/')

# Вход в акаунт
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="i0116"]'))).send_keys(credentials['login'])


def dalie():
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="idSIButton9"]'))).click()


dalie()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="i0118"]'))).send_keys(credentials['password'])

dalie()


WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="idBtn_Back"]'))).click()


print("------------------------Вход в акаунт прошёл успешно!------------------------")

# Переход на почту  и выбор собщения

driver.get('https://outlook.live.com/mail/')


def download_button():  # Скачка собщения
    print("Скачка собщения----------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@name = "Download"]'))).click()
    print("------------------------Скачивание  собщения прошло успешно!------------------------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[@title = "Close" ]'))).click()


# Скач# Скачка А
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training A"]'))).click()
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title = "Training A.xlsx"]'))).click()
print("------------------------Выбор собщения прошёл успешно!------------------------")
download_button()
# Скачка B
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training B"]'))).click()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title = "Training B.xlsx"]'))).click()

download_button()


WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training C"]'))).click()
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title = "Training C.xlsx"]'))).click()

download_button()

# Парсинг собщения из файл
print("------------------------Парсинг  собщения!------------------------")

list_bad = []
for f in [os.getcwd() + '/downloaded_files/Training A.xlsx',
          os.getcwd() + '/downloaded_files/Training B.xlsx',
          os.getcwd() + '/downloaded_files/Training C.xlsx']:
    data = pd.read_excel(f)
    mail = data['Mail'].tolist()
    list_bad.extend(mail)
    mail_list = list(set(list_bad))
print("------------------------Парсинг  собщения прошёл успешно!------------------------")
for email in mail_list:
    print("-----------------Создание Собщения!-----------------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[text()="New message"]'))).click()
    print("-----------------Добавление юзеров!-----------------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@aria-label = "To"]'))).send_keys(email)

    print("-----------------СозданиЕ Темы!-----------------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@aria-label = "Add a subject"]'))).send_keys('You need to pass a training')
    print("-----------------Создание самого собщения!-----------------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@aria-label="Message body"]'))).send_keys('Hello! You need to pass trainings. Have a nice day!')
    print("-----------------Отправка Собщения!-----------------")
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[contains(@title, "Send")]'))).click()
    print("----------------- Собщение Отправлено!-----------------")

sleep(5)
driver.close()
driver.quit()
