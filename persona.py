import time
import tkinter as tk
import sys
import os
from datetime import datetime, date
from getpass import getuser
from icalendar import Calendar, Event
from loguru import logger
from pandas import DataFrame
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from timeit import default_timer as timer

def check_platform():
    platform = sys.platform
    user = getuser()
    if platform == 'win32':
        profile_path = f'C:\\Users\\{user}\\AppData\\Local\\Google\\Chrome\\User Data'
        path = os.getcwd() + '\\chromedriver.exe'
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        return profile_path, path, desktop
    else:
        profile_path = f'/Users/{user}/Library/Application Support/Google/Chrome/'
        path = os.getcwd() + '/chromedriver'
        desktop = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
        return profile_path, path, desktop

def incr_month():
    month = today.strftime('%m')
    x = int(month)
    x += 1
    month = str(x)
    return '/' + month

def strip_chars(date):
    sep = '\n'
    integer = int(date.split(sep, 1)[0])
    return integer

def scrape(month):
    logger.debug('Scraping data')
    x = 2
    # year = today.strftime('/%Y')
    while x < 9:
        try:
            shift_date = driver.find_element_by_xpath(
                f'//*[@id="headercontainer-1037-targetEl"]/div[{x}]/div/span/span'
            ).text
            shift_day = strip_chars(shift_date)
            if len(int_days) < 1 or shift_day > int_days[-1]:
                int_days.append(shift_day)
                # days.append(str(shift_day) + month + year)
            else:
                month = incr_month()
                # days.append(str(shift_day) + month + year)
            shift_times = driver.find_element_by_xpath(
                f'//*[@id="gridview-1046-record-scheduledHoursRow"]/td[{x}]/div/div/div/div'
            ).text
            start.append(shift_times[:5])
            end.append(shift_times[-5:])
            shift_hours = driver.find_element_by_xpath(
                f'//*[@id="gridview-1046-record-scheduledHoursRow"]/td[{x}]/div/div/div[2]/div'
            ).text
            hours.append('Work - ' + shift_hours[:-3] + 'hours')
            x += 1
        except NoSuchElementException:
            int_days.pop()
            # days.pop()
            x += 1
    return month

def go_next():
    action = ActionChains(driver)
    action.click(
        driver.find_element_by_xpath('//*[@id="button-1029-btnIconEl"]'))
    action.perform()
    logger.debug('Accessing next page')

def check_if_working():
    if driver.find_element_by_xpath(
            '//*[@id="panel-1032-innerCt"]/div/div[4]/span/div/div'
    ).text == 'Schedules have not yet been posted.' or driver.find_element_by_xpath(
            '//*[@id="gridview-1046-body"]/tr[1]/td/div/div[3]'
    ).text == '0.00 hrs':
        logger.debug('No data found')
        return False
    else:
        logger.debug('Data found')
        return True

def export_data(month):
    csv['summary'] = hours
    # csv['Start Date'] = days
    # csv['Start Time'] = start
    # csv['End Date'] = days
    # csv['End Time'] = end
    df = DataFrame(csv)
    # df.to_csv(desktop + '/rota.csv', index=None, header=True)
    print(df)

    cal = Calendar()
    # month = int(today.strftime('%m'))
    year = int(today.strftime('%Y'))
    month = int(month.strip('/'))

    data = list(df['summary'])

    i = 0
    for column in data:
        event = Event()
        event.add('summary', df['summary'][i])
        event.add('dtstart', datetime(year, month, int_days[i], int(start[i][:2]), int(start[i][3:])))
        event.add('dtend', datetime(year, month, int_days[i], int(end[i][:2]), int(end[i][3:])))
        event.add('location', 'Microsoft Store - Oxford Circus, 253-259 Regent St, Mayfair, London W1B 2ER, UK')
        cal.add_component(event)
        i += 1

    with open(os.path.join(desktop, 'persona_calendar.ics'), 'wb') as f:
        f.write(cal.to_ical())


root = tk.Tk()

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

today = date.today()

int_days = []

# subject = []
# days = []
hours = []
start = []
end = []
csv = {
    'summary': [],
    # 'Start Date': [],
    # 'Start Time': [],
    # 'End Date': [],
    # 'End Time': [],
}

profile_path, path, desktop = check_platform()
options = Options()
options.add_argument(f'window-size={screen_width}x{screen_height}')
options.add_argument('user-data-dir=' + profile_path) # use custom chrome profile with 'keep signed in' on persona saved 
options.headless = True
with Chrome(path, options=options) as driver:
    t_start = timer()
    driver.implicitly_wait(2)
    try:
        driver.get('https://aka.ms/personaeu')
        logger.debug('Accessing Persona')
        driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))
    except:
        sys.exit('Could not access persona, script exited.')

    working = check_if_working()
    month = str(today.strftime('/%m'))
    while working == True:
        month = scrape(month)
        go_next()
        time.sleep(1)
        working = check_if_working()
export_data(month)

t_end = timer()
logger.debug('\nRan script in ' + str(t_end - t_start) + 's')
