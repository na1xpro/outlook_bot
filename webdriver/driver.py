from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from sys import platform
import os
from webdriver_manager.chrome import ChromeDriverManager

if platform == 'win32':
    plat = r'\downloaded_files'
else:
    plat = r'/downloaded_files'

options = Options()
options.add_argument('--window-size=1400,800')
options.add_experimental_option(
    'prefs',
    {
        'download.default_directory': os.getcwd() + plat,
        'download.prompt_for_download': False,
        'safebrowsing.enabled': True,
        'download.directory_upgrade': True,
    },
)

driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
