import os
from time import sleep
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import glob
from webdriver.driver import driver
from loguru import logger
from sys import platform
from constants import credentials
from constants import put_message
from constants import os_path

driver.get('https://login.live.com/')

# Вход в акаунт
WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="i0116"]'))).send_keys(
    credentials['login'])


def dalie():
    WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="idSIButton9"]'))).click()


dalie()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="i0118"]'))).send_keys(
    credentials['password'])

dalie()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//input[@id="idBtn_Back"]'))).click()

logger.info("Account authorization successful.")

# Переход на почту  и выбор собщения

driver.get('https://outlook.live.com/mail/')

logger.info('Check for the presence of an existing file in a folder.')
logger.warning('The files in the folder may be deleted!')
pas = os.getcwd() + os_path[platform]['path_download_folder']
files = glob.glob(pas + r'\*.xlsx')
for filename in files:
    os.unlink(filename)


def download_button():
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, '//div[@class = "_24WqHp8mfxSp2QIJMkmSrM"]'))).click()
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, "//button[@title = 'More actions']"))).click()
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, "//i[@data-icon-name = 'Download']"))).click()


WebDriverWait(driver, 13).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training A"]'))).click()

download_button()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training B"]'))).click()

download_button()

WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, '//div[@title="Training C"]'))).click()

download_button()

logger.info('Parsing a message...')
list_bad = []
os.chdir(os.getcwd())
for file in files:
    data = pd.read_excel(file)
    mail = data['Mail'].tolist()
    list_bad.extend(mail)
    mail_list = list(set(list_bad))


logger.info('Parsing of the message was successful!')

for email in mail_list:
    logger.info("Create a message!")
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, '//span[text()="New message"]'))).click()
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, '//input[@aria-label = "To"]'))).send_keys(email)
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, '//input[@aria-label = "Add a subject"]'))).send_keys(
        put_message['title_message'])
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, '//div[@aria-label="Message body"]'))).send_keys(
        put_message['message'])
    WebDriverWait(driver, 10).until(
        ec.visibility_of_element_located((By.XPATH, '//button[contains(@title, "Send")]'))).click()
    logger.info('Message sent!')

sleep(3)
driver.close()
driver.quit()
