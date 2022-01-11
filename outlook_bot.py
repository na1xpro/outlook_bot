import os
from time import sleep
import pandas as pd

from selenium.webdriver.common.by import By

from webdriver.driver import driver
from constants import credentials

driver.get('https://login.live.com/')

# Вход в акаунт
driver.find_element(By.NAME, "loginfmt").send_keys(credentials.get['login'])


def dalie():
    driver.find_element(By.ID, "idSIButton9").click()


dalie()
sleep(1)
password = driver.find_element(By.NAME, "passwd").send_keys(credentials['password'])
sleep(1)
dalie()
# Тест строка
driver.find_element(By.ID, "KmsiCheckboxField").click()

driver.find_element(By.ID, "idBtn_Back").click()
print("------------------------Вход в акаунт прошёл успешно!------------------------")

# Переход на почту  и выбор собщения

driver.get('https://outlook.live.com/mail/')
sleep(4)


def download_button():  # Скачка собщения
    sleep(5)

    print("Скачка собщения----------")
    driver.find_element(By.XPATH, '//button[@name = "Download"]').click()
    print("------------------------Скачивание  собщения прошло успешно!------------------------")
    driver.find_element(By.XPATH, '//button[@title = "Close" ]').click()


driver.find_element(By.XPATH, '//div[@title="Training A"]').click()  # Скачка А
sleep(3)
driver.find_element(By.XPATH, '//div[@title = "Training A.xlsx"]').click()
print("------------------------Выбор собщения прошёл успешно!------------------------")
download_button()

driver.find_element(By.XPATH, '//div[@title="Training B"]').click()  # Скачка B
sleep(3)
driver.find_element(By.XPATH, '//div[@title = "Training B.xlsx"]').click()
download_button()

driver.find_element(By.XPATH, '//div[@title="Training C"]').click()  # Скачка C
sleep(3)
enter_message = driver.find_element(By.XPATH, '//div[@title = "Training C.xlsx"]').click()
download_button()

# Парсинг собщения из файл
print("------------------------Парсинг  собщения!------------------------")
sleep(3)
list_bad = []
for f in [os.getcwd() + '/downloaded_files/Training A.xlsx',
          os.getcwd() + '/downloaded_files/Training B.xlsx',
          os.getcwd() + '/downloaded_files/Training C.xlsx']:
    data = pd.read_excel(f)
    mail = data['Mail'].tolist()
    list_bad.extend(mail)
    mail_list = list(set(list_bad))
print("------------------------Парсинг  собщения прошёл успешно!------------------------")

sleep(4)

for email in mail_list:
    print("-----------------Создание Собщения!-----------------")
    open_message_icon = driver.find_element(By.XPATH, '//span[text()="New message"]').click()
    sleep(5)
    print("-----------------Добавление юзеров!-----------------")
    users_message = driver.find_element(By.XPATH, '//input[@aria-label = "To"]').send_keys(email)

    sleep(3)
    print("-----------------СозданиЕ Темы!-----------------")
    title_message = driver.find_element(By.XPATH, '//input[@aria-label = "Add a subject"]')

    title_message.send_keys('You need to pass a training')

    print("-----------------Создание самого собщения!-----------------")
    message_for_users = driver.find_element(By.XPATH, '//div[@aria-label="Message body"]').send_keys(
        'Hello! You need to pass trainings. Have a nice day!')
    print("-----------------Отправка Собщения!-----------------")
    message_send = driver.find_element(By.XPATH, '//button[contains(@title, "Send")]').click()
    print("----------------- Собщение Отправлено!-----------------")

sleep(5)
driver.close()
driver.quit()
