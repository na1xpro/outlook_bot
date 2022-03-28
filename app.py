from outlook_bot import Bot
from webdriver.driver import driver
from webdriver.driver import plat

bot = Bot(plat, driver)
bot.auth()
bot.checking_files_in_folder()
bot.download_files()
bot.parsing_message()
bot.send_message()

driver.close()
driver.quit()
