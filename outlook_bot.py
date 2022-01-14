import os
from time import sleep
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from webdriver.driver import driver
from loguru import logger
import glob
from constants import credentials
from constants import put_message


driver.get('https://login.live.com/')

# Вход в акаунт
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="i0116"]'))).send_keys(credentials['login'])


def dalie():
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="idSIButton9"]'))).click()


dalie()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="i0118"]'))).send_keys(credentials['password'])

dalie()


WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="idBtn_Back"]'))).click()


logger.info("Account authorization successful.")

# Переход на почту  и выбор собщения

driver.get('https://outlook.live.com/mail/')

logger.info('Check for the presence of an existing file in a folder.')


# Проверка на наличие файла
os.chdir(os.getcwd() + "/downloaded_files")
logger.warning("If it is found it will be deleted!")
files = glob.glob('*.xlsx')
for filename in files:
    os.unlink(filename)


def download_button():  # Скачка собщения
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@class = "_24WqHp8mfxSp2QIJMkmSrM"]'))).click()
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, "//button[@title = 'More actions']"))).click()
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, "//i[@data-icon-name = 'Download']"))).click()


WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training A"]'))).click()

download_button()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training B"]'))).click()

download_button()


WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training C"]'))).click()

download_button()

logger.info('Parsing a message...')
list_bad = []
os.chdir(os.getcwd())
for file in glob.glob("*.xlsx"):
    data = pd.read_excel(file)
    mail = data['Mail'].tolist()
    list_bad.extend(mail)
    mail_list = list(set(list_bad))
logger.info('Parsing of the message was successful!')

for email in mail_list:
    logger.info('Create a message!')
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//span[text()="New message"]'))).click()
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@aria-label = "To"]'))).send_keys(email)
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@aria-label = "Add a subject"]'))).send_keys(put_message['title_message'])
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@aria-label="Message body"]'))).send_keys(put_message['message'])
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//button[contains(@title, "Send")]'))).click()
    logger.info('Message sent!')

sleep(3)
driver.close()
driver.quit()
