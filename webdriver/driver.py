from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from sys import platform
import os

if "linux" in platform:
    path_driver = '/webdriver/chromedriver'
elif platform == "win32":
    path_driver = r'\webdriver\chromedriver.exe'

way = os.getcwd() + path_driver
os.chmod(way, 755)
driver_service = Service(way,)

options = Options()
options.add_argument('--window-size=1400,800')
options.add_experimental_option('prefs', {
    'download.default_directory': os.getcwd() + r'\downloaded_files',
    'download.prompt_for_download': False,
    'safebrowsing.enabled': True,
    'download.directory_upgrade': True,
})

driver = webdriver.Chrome(service=driver_service, options=options)
