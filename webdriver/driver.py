from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from sys import platform
import os

from constants import os_path

path_driver = os.getcwd() + os_path[platform]['path_driver']
os.chmod(path_driver, 755)
driver_service = Service(path_driver)

options = Options()
options.add_argument('--window-size=1400,800')
options.add_experimental_option(
    'prefs',
    {
        'download.default_directory': os.getcwd() + os_path[platform]['path_download_folder'],
        'download.prompt_for_download': False,
        'safebrowsing.enabled': True,
        'download.directory_upgrade': True,
    },
)

driver = webdriver.Chrome(service=driver_service, options=options)