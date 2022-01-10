from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import os
from sys import platform



if platform == "linux" or platform == "linux2":
    way_lin = os.getcwd() + '/webdriver/chromedriver'
    os.chmod(way_lin, 755)
    driver_service = Service(way_lin)


elif platform == "win32":
    way_win = os.getcwd() + '\chromedriver.exe'
    os.chmod(way_win, 755)
    driver_service = Service(way_win)


options = Options()
options.add_argument('--window-size=1400,800')
options.add_experimental_option('prefs', {
    'download.default_directory': os.getcwd() + '\downloaded_files',
    'download.prompt_for_download': False,
    'safebrowsing.enabled': True,
    'download.directory_upgrade': True,
})

driver = webdriver.Chrome(service=driver_service, options=options)
