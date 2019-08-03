import sys
import os
from datetime import date
from pandas import DataFrame
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

today = date.today()
date = today.strftime('/%m/%Y')

platform = sys.platform
if platform == 'win32':
    profile_path = 'C:\\Users\\AVTORRES\\AppData\\Local\\Google\\Chrome\\User Data'
    path = os.getcwd() + '\\chromedriver.exe'
    desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
else:
    profile_path = '/Users/arlandtorres/Library/Application Support/Google/Chrome/'
    path = os.getcwd() + '/chromedriver'
    desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')

options = Options()
options.add_argument('user-data-dir=' + profile_path)
# options.add_argument('headless')

driver = webdriver.Chrome(path, options=options)
driver.implicitly_wait(4)
print('accessing persona...')
driver.get('https://aka.ms/personaeu')
driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))

days = []
hours = []
start = []
end = []
csv = {
    'Subject': [],
    'Start Date': [],
    'Start Time': [],
    'End Date': [],
    'End Time': [],
    'Reminder Date': [],
}

def scrape():
    x = 2
    while x < 9:
        try:
            shift_days = driver.find_element_by_xpath(
                f'//*[@id="headercontainer-1037-targetEl"]/div[{x}]/div/span/span'
            ).text
            days.append(shift_days[:2] + date)
            shift_times = driver.find_element_by_xpath(
                f'//*[@id="gridview-1046-record-scheduledHoursRow"]/td[{x}]/div/div/div/div'
            ).text
            start.append(shift_times[:5])
            end.append(shift_times[-5:])
            shift_hours = driver.find_element_by_xpath(
                f'//*[@id="gridview-1046-record-scheduledHoursRow"]/td[{x}]/div/div/div[2]/div'
            ).text
            hours.append('WORK - ' + shift_hours)
            x += 1
        except NoSuchElementException:
            days.pop()
            # start.append('OFF')
            # end.append('OFF')
            # hours.append('OFF')
            x += 1

def go_next():
    action = ActionChains(driver)
    action.click(driver.find_element_by_xpath('//*[@id="button-1029-btnIconEl"]'))
    action.perform()

def check_if_working():
    if driver.find_element_by_xpath('//*[@id="gridview-1046-body"]/tr[1]/td/div/div[3]').text != '0.00 hrs':
        print('you\'re working next week buddy...')
        return True
    else:
        print('you have nothing scheduled next week...')
        return False

def create_csv():
    df = DataFrame(csv)
    # df.to_csv(desktop + '/rota.csv', index=None, header=True)
    print('here is your schedule :)')
    print(df)
    
working = check_if_working()
while working == True:
    print('scraping data...')
    scrape()
    print('checking next week...')
    go_next()
    working = check_if_working()
driver.quit()

days = [s.replace('\n', '') for s in days]
csv['Subject'] = hours
csv['Start Date'] = days
csv['Start Time'] = start
csv['End Date'] = days
csv['End Time'] = end
csv['Reminder Date'] = days
create_csv()