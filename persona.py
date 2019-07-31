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


def incr_month():
    month = today.strftime('%m')
    x = int(month)
    x += 1
    month = str(x)
    return '/' + month


def scrape(month):
    x = 2
    year = today.strftime('/%Y')
    while x < 9:
        try:
            shift_days = driver.find_element_by_xpath(
                f'//*[@id="headercontainer-1037-targetEl"]/div[{x}]/div/span/span'
            ).text
            if len(days) < 1:
                days.append(shift_days[:2] + month + year)
            elif shift_days[:2] > days[-1][:2]:
                days.append(shift_days[:2] + month + year)
            else:
                month = incr_month()
                days.append(shift_days[:2] + month + year)
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
            x += 1
    return month


def go_next():
    action = ActionChains(driver)
    action.click(
        driver.find_element_by_xpath('//*[@id="button-1029-btnIconEl"]'))
    action.perform()


def check_if_working():
    if driver.find_element_by_xpath(
            '//*[@id="panel-1032-innerCt"]/div/div[4]/span/div/div'
    ).text == 'Schedules have not yet been posted.' or driver.find_element_by_xpath(
            '//*[@id="gridview-1046-body"]/tr[1]/td/div/div[3]'
    ).text == '0.00 hrs':
        return False
    else:
        return True


def create_csv():
    df = DataFrame(csv)
    df.to_csv(desktop + '/rota.csv', index=None, header=True)
    print(df)


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
options.add_argument('headless')

driver = webdriver.Chrome(path, options=options)
driver.implicitly_wait(2)
driver.get('https://aka.ms/personaeu')
driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))

working = check_if_working()
month = str(today.strftime('/%m'))
while working == True:
    month = scrape(month)
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
