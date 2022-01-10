from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os

driver_path = Service(os.getcwd() + '\chromedriver.exe')

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