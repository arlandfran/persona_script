import sys
import os
from datetime import date

from selenium import webdriver

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException

today = date.today()

month = today.strftime('%m')

options = Options()

platform = sys.platform

if platform == 'win32':
    profile_path = 'C:\\Users\\AVTORRES\\AppData\\Local\\Google\\Chrome\\User Data'
    path = os.getcwd() + '\\chromedriver.exe'
else:
    profile_path = '/Users/arlandtorres/Library/Application Support/Google/Chrome/'
    path = os.getcwd() + '/chromedriver'

options.add_argument('user-data-dir=' + profile_path)

driver = webdriver.Chrome(path, options=options)

driver.implicitly_wait(3)

driver.get('https://aka.ms/personaeu')

driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))

# days = []

# schedule = []

# hours = []

# x = 2

# while x < 9:
#     try:
#         shift_days = driver.find_element_by_xpath(
#             f'//*[@id="headercontainer-1037-targetEl"]/div[{x}]/div/span/span'
#         ).text
#         days.append(shift_days)
#         shift_times = driver.find_element_by_xpath(
#             f'//*[@id="gridview-1046-record-scheduledHoursRow"]/td[{x}]/div/div/div/div'
#         ).text
#         schedule.append(shift_times)
#         shift_hours = driver.find_element_by_xpath(
#             f'//*[@id="gridview-1046-record-scheduledHoursRow"]/td[{x}]/div/div/div[2]/div'
#         ).text
#         hours.append(shift_hours)
#         x += 1
#     except NoSuchElementException:
#         schedule.append('OFF')
#         hours.append('OFF')
#         x += 1

action = ActionChains(driver)
action.click(driver.find_element_by_xpath('//*[@id="button-1029-btnIconEl"]'))
action.perform()

# driver.quit()