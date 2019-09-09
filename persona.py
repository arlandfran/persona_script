import time
import tkinter as tk
import sys
import os
from datetime import datetime, date
from getpass import getuser
from icalendar import Calendar, Event
from loguru import logger
# from PyInquirer import prompt
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver import Chrome
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from tempfile import mkdtemp
from timeit import default_timer as timer


def check_platform():
    platform = sys.platform
    user = getuser()
    if platform == 'win32':
        profile_path = f'C:\\Users\\{user}\\AppData\\Local\\Google\\Chrome\\User Data'
        path = os.getcwd() + '\\chromedriver.exe'
        desktop = os.path.join(os.path.join(
            os.environ['USERPROFILE']), 'Desktop')
        return profile_path, path, desktop
    else:
        profile_path = f'/Users/{user}/Library/Application Support/Google/Chrome/'
        path = os.getcwd() + '/chromedriver'
        desktop = os.path.join(os.path.join(
            os.path.expanduser('~')), 'Desktop')
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
    while x < 9:
        try:
            shift_date = driver.find_element_by_xpath(
                f'//*[@id="headercontainer-1037-targetEl"]/div[{x}]/div/span/span'
            ).text
            shift_day = strip_chars(shift_date)
            if len(int_days) < 1 or shift_day > int_days[-1]:
                int_days.append(shift_day)
            else:
                month = incr_month()
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
    cal = Calendar()
    year = int(today.strftime('%Y'))
    month = int(month.strip('/'))
    location = 'Microsoft Store - Oxford Circus, 253-259 Regent St, Mayfair, London W1B 2ER, UK'
    i = 0
    for item in hours:
        event = Event()
        event.add('summary', hours[i])
        event.add('dtstart', datetime(
            year, month, int_days[i], int(start[i][:2]), int(start[i][3:])))
        event.add('dtend', datetime(
            year, month, int_days[i], int(end[i][:2]), int(end[i][3:])))
        event.add('location', location)
        cal.add_component(event)
        i += 1

    temp_dir = mkdtemp()
    file_path = os.path.join(temp_dir, 'persona_calendar.ics')
    with open(file_path, 'wb') as f:
        f.write(cal.to_ical())
    return file_path


def outlook_import(path):
    driver.get('https://www.outlook.com/calendar')
    driver.find_element_by_xpath(
        '//*[@id="app"]/div/div[2]/div[1]/div/div[1]/div[3]/div/div[1]/button[2]').click()
    driver.find_element_by_xpath('//*[@id="Fromfile"]').click()
    driver.find_element_by_xpath(
        '//input[@type="file"]').send_keys(path)
    driver.find_element_by_xpath(
        '/html/body/div[6]/div/div/div/div[2]/div[2]/div/div[2]/div/div[2]/div[6]/button').click()
    logger.debug('Calendar imported into Outlook')


def gcalendar_import(path):
    driver.get('https://calendar.google.com')
    driver.find_element_by_xpath(
        '/html/body/div[2]/div[1]/div[1]/div[2]/div[1]/div/div[2]/div[4]/div/div/div/div/div[4]/div[2]/div/div[2]').click()
    driver.find_element_by_xpath('/html/body/div[21]/div/div/span[5]').click()
    driver.find_element_by_xpath(
        '//*[@id="YPCqFe"]/div/div/div/div[1]/div[1]/div/form/label/input').send_keys(path)
    driver.find_element_by_xpath(
        '//*[@id="YPCqFe"]/div/div/div/div[1]/div[3]/div[2]').click()
    logger.debug('Calendar imported into Google Calendar')


def import_prompt():
    # question = [
    #     {
    #         'type': 'list',
    #         'message': 'Select calendar:',
    #         'name': 'Calendar',
    #         'choices': [
    #             {
    #                 'name': 'Outlook'
    #             },
    #             {
    #                 'name': 'Google Calendar'
    #             },
    #         ]
    #     }
    # ]

    # answer = prompt(question)
    # return answer

    answer = input(
        'Please select a calendar to import to: (Enter the corresponding number to select)\n1. Outlook\n2. Google Calendar\n\n')
    return answer


root = tk.Tk()
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

today = date.today()

int_days = []
hours = []
start = []
end = []
csv = {'summary': []}

profile_path, path, desktop = check_platform()
options = Options()
# options.add_argument("--start-maximized")
options.add_argument(f'window-size={screen_width}x{screen_height}')
# use custom chrome profile with 'keep signed in' on persona saved
options.add_argument('user-data-dir=' + profile_path)
options.headless = True
selection = import_prompt()
with Chrome(path, options=options) as driver:
    t_start = timer()
    driver.implicitly_wait(4)
    try:
        driver.get('https://aka.ms/personaeu')
        logger.debug('Accessing Persona')
        driver.switch_to.frame(driver.find_element_by_tag_name('iframe'))
    except:
        sys.exit('Could not access persona, script exited.')

    working = check_if_working()
    month = str(today.strftime('/%m'))
    while working:
        month = scrape(month)
        go_next()
        time.sleep(1)
        working = check_if_working()
    file_path = export_data(month)
    if selection == '1':
        outlook_import(file_path)
    elif selection == '2':
        gcalendar_import(file_path)

t_end = timer()
logger.debug('Ran script in ' + str(t_end - t_start) + 's')
