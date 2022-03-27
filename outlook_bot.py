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
class Bot:
    def __init__(self, plat, driver, ):
        self.plat = plat
        self.driver = driver
        self.get_giles()

    def get_giles(self):
        pas = os.getcwd() + self.plat
        files = glob.glob(pas + r'/*.xlsx')
        return files

    def find_element(self, xpath):
        return WebDriverWait(self.driver, 10).until(ec.visibility_of_element_located((By.XPATH, xpath)))

    # Вход в акаунт
    def auth(self):
        self.find_element('//input[@name = "loginfmt"]').send_keys(credentials['login'])
        submit = '//input[@type = "submit"]'
        self.find_element(submit).click()
        self.find_element('//input[@name = "passwd"]').send_keys(credentials['password'])
        self.find_element(submit).click()
        self.find_element('//input[@id="idBtn_Back"]').click()
        logger.info("Account authorization successful.")
        driver.get('https://outlook.live.com/mail/')

    def checking_files_in_folder(self):
        logger.info('Check for the presence of an existing file in a folder.')
        logger.warning('The files in the folder may be deleted!')
        for filename in files_list:
            os.unlink(filename)

    def download_files(self):

        folders_file = ("Training A", "Training B", "Training C")
        for i in folders_file:
            self.find_element('//div[@title=' + '"' + i + '"' + ']').click()
            sleep(1)
            self.find_element('//div[@title = "Select all messages"]').click()
            sleep(1)
            self.find_element("//button[@title = 'More actions']").click()
            sleep(1)
            self.find_element("//i[@data-icon-name = 'Download']").click()

    #    Парсинг собщенимя
    def parsing_message(self, files_list):
        logger.info('Parsing a message...')
        list_bad = []
        for file in files_list:
            data = pd.read_excel(file)
            mail = data['Mail'].tolist()
            list_bad.extend(mail)
            mail_list = list(set(list_bad))
            return mail_list

        logger.info('Parsing of the message was successful!')

    def send_message(self, mail_list):
        for email in mail_list:
            logger.info("Create a message!")
            self.find_element('//span[text()="New message"]').click()
            self.find_element('//input[@aria-label = "To"]').send_keys(email)
            self.find_element('//input[@aria-label = "Add a subject"]').send_keys(put_message['title_message'])
            self.find_element('//div[@aria-label="Message body"]').send_keys(put_message['message'])
            self.find_element('//button[contains(@title, "Send")]').click()
            logger.info('Message sent!')


bot = Bot(plat, driver, )
bot.auth()
bot.checking_files_in_folder()
bot.download_files()

# auth()
# checking_files_in_folder(get_giles())
# download_files()
# mails = parsing_message(get_giles())
# send_message(mails)
# sleep(3)
# driver.close()
# driver.quit()
