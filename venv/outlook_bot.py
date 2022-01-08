from selenium import webdriver
from time import sleep
from selenium.webdriver.chrome.options import Options
import pandas as pd
import os.path
import warnings

warnings.filterwarnings("ignore") #Отключение варнингов
# Перепатч хрома на загрузку файлов в нужную нам папку, а именно в корневой каталог проекта в папку downloaded_files

driver_path = os.getcwd() + '/webdriver/chromedriver.exe'

options = Options()

options.add_experimental_option('prefs', {
    'download.default_directory': os.getcwd() + '\downloaded_files',
    'download.prompt_for_download': False,
    'safebrowsing.enabled': True,
    'download.directory_upgrade': True,
})

driver_login = webdriver.Chrome(executable_path=driver_path, chrome_options=options)

driver_login.get('https://login.live.com/')

# Вход в акаунт
email = driver_login.find_element_by_name("loginfmt")
email.send_keys("vivexpro@outlook.com")


def dalie():
    knopka_dalie = driver_login.find_element_by_id("idSIButton9").click()


dalie()
sleep(1)
password = driver_login.find_element_by_name("passwd")
password.send_keys("Vivexpass1")
sleep(1)
dalie()

knopka_bolshe_ne_pokazivat = driver_login.find_element_by_id("KmsiCheckboxField").click()

knopka_net = driver_login.find_element_by_id("idBtn_Back").click()
print("------------------------Вход в акаунт прошлёл успешно!------------------------")

# Переход на почту  и выбор собщения

driver_login.get('https://outlook.live.com/mail/')
sleep(4)


def download_button():  # Скачка собщения
    sleep(5)

    print("Скачка собщения----------")
    download_message = driver_login.find_element_by_xpath('//button[@name = "Download"]').click()
    print("------------------------Скачивание  собщения прошло успешно!------------------------")
    close_form_download = driver_login.find_element_by_xpath('//button[@title = "Close" ]').click()


enter_trainingA = driver_login.find_element_by_xpath('//div[@title="Training A"]').click()  # Скачка А
sleep(3)
enter_message = driver_login.find_element_by_xpath('//div[@title = "Training A.xlsx"]').click()
print("------------------------Выбор собщения прошёл успешно!------------------------")
download_button()

enter_trainingB = driver_login.find_element_by_xpath('//div[@title="Training B"]').click()  # Скачка B
sleep(3)
enter_message = driver_login.find_element_by_xpath('//div[@title = "Training B.xlsx"]').click()
download_button()

enter_trainingC = driver_login.find_element_by_xpath('//div[@title="Training C"]').click()  # Скачка C
sleep(3)
enter_message = driver_login.find_element_by_xpath('//div[@title = "Training C.xlsx"]').click()
download_button()

# Парсинг собщения из файл
print("------------------------Парсинг  собщения!------------------------")
sleep(3)
list_bad = []
for f in [os.getcwd() + '/downloaded_files/Training A.xlsx',
          os.getcwd() + '/downloaded_files/Training B.xlsx',
          os.getcwd() + '/downloaded_files/Training C.xlsx']:
    data = pd.read_excel(f, )
    mail = data['Mail'].tolist()
    list_bad.extend(mail)
    mail_list = list(set(list_bad))
print("------------------------Парсинг  собщения прошёл успешно!------------------------")

sleep(4)

for email in mail_list:
    print("-----------------СозданиЕ Собщения!-----------------")
    open_message_icon = driver_login.find_element_by_xpath('//span[text()="New message"]').click()
    sleep(5)
    print("-----------------Добавление юзеров!-----------------")
    users_message = driver_login.find_element_by_xpath('//input[@aria-label = "To"]')
    users_message.send_keys(email)

    sleep(3)
    print("-----------------СозданиЕ Темы!-----------------")
    title_message = driver_login.find_element_by_xpath('//input[@aria-label = "Add a subject"]')

    title_message.send_keys('You need to pass a training')

    print("-----------------СозданиЕ Самого собщения!-----------------")
    message_for_users = driver_login.find_element_by_xpath('//div[@aria-label="Message body"]')
    message_for_users.send_keys('Hello! You need to pass trainings. Have a nice day!')

    print("-----------------Отправка Собщения!-----------------")
    message_send = driver_login.find_element_by_xpath('//button[contains(@title, "Send")]').click()
    print("----------------- Собщение Отправлено!-----------------")

sleep(5)
driver_login.close()
driver_login.quit()
