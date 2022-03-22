import os
from time import sleep
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import glob
from webdriver.driver import driver
from loguru import logger
from constants import credentials, put_message
from webdriver.driver import plat

driver.get('https://login.live.com/')


#  Хуя тут э смалыки 🙃
def get_giles():
    pas = os.getcwd() + plat
    files = glob.glob(pas + r'/*.xlsx')
    return files


def find_element(xpath):
    return WebDriverWait(driver, 10).until(ec.visibility_of_element_located((By.XPATH, xpath)))


# Вход в акаунт
def auth():
    find_element('//input[@name = "loginfmt"]').send_keys(credentials['login'])
    submit = '//input[@type = "submit"]'
    find_element(submit).click()
    find_element('//input[@name = "passwd"]').send_keys(credentials['password'])
    find_element(submit).click()
    find_element('//input[@id="idBtn_Back"]').click()
    logger.info("Account authorization successful.")
    driver.get('https://outlook.live.com/mail/')


def checking_files_in_folder(files_list):
    logger.info('Check for the presence of an existing file in a folder.')
    logger.warning('The files in the folder may be deleted!')
    for filename in files_list:
        os.unlink(filename)


def download_files():
    def download_button():
        sleep(1)
        find_element('//div[@title = "Select all messages"]').click()
        sleep(1)
        find_element("//button[@title = 'More actions']").click()
        sleep(1)
        find_element("//i[@data-icon-name = 'Download']").click()

    folders_file = ("Training A", "Training B", "Training C")
    for i in folders_file:
        find_element('//div[@title=' + '"' + i + '"' + ']').click()
        download_button()


# Парсинг собщенимя
def parsing_message(files_list):
    logger.info('Parsing a message...')
    list_bad = []
    for file in files_list:
        data = pd.read_excel(file)
        mail = data['Mail'].tolist()
        list_bad.extend(mail)
        mail_list = list(set(list_bad))
        return mail_list

    logger.info('Parsing of the message was successful!')


def send_message(mail_list):
    for email in mail_list:
        logger.info("Create a message!")
        find_element('//span[text()="New message"]').click()
        find_element('//input[@aria-label = "To"]').send_keys(email)
        find_element('//input[@aria-label = "Add a subject"]').send_keys(put_message['title_message'])
        find_element('//div[@aria-label="Message body"]').send_keys(put_message['message'])
        find_element('//button[contains(@title, "Send")]').click()
        logger.info('Message sent!')


auth()
checking_files_in_folder(get_giles())
download_files()
mails = parsing_message(get_giles())
send_message(mails)
sleep(3)
driver.close()
driver.quit()
