from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os
from sys import platform



if platform == "linux" or platform == "linux2":
    driver_path = Service(os.getcwd() + '/webdriver/chromedriver')

elif platform == "win32":
    driver_path = Service(os.getcwd() + '\webdriver\chromedriver.exe')



options = Options()
options.add_argument('--window-size=1400,800')
options.add_experimental_option('prefs', {
    'download.default_directory': os.getcwd() + '\downloaded_files',
    'download.prompt_for_download': False,
    'safebrowsing.enabled': True,
    'download.directory_upgrade': True,
})

driver = webdriver.Chrome(service=driver_path, options=options)
print('path:', os.getcwd())
