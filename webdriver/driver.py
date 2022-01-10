from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os
from sys import platform



if platform == "linux" or platform == "linux2":
    way = os.getcwd() + '/webdriver/chromedriver'
    os.chmod(way, 755)
    driver_service = Service(way)


elif platform == "win32":
    way = os.getcwd() + '\webdriver\chromedriver.exe'
    os.chmod(way, 755)
    driver_service = Service(way)


options = Options()
options.add_argument('--window-size=1400,800')
options.add_experimental_option('prefs', {
    'download.default_directory': os.getcwd() + '\downloaded_files',
    'download.prompt_for_download': False,
    'safebrowsing.enabled': True,
    'download.directory_upgrade': True,
})

driver = webdriver.Chrome(service=driver_service, options=options)
