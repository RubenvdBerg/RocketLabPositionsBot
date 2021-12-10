import time
from datetime import datetime
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import os
from selenium.webdriver.support.ui import Select

driver = webdriver.Chrome('chromedriver/chromedriver.exe')
driver.get("https://www.rocketlabusa.com/careers/positions/")
location_dropdown = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[1]/div/div/div')
location_dropdown.click()
time.sleep(.5)
auckland_option = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[1]/div/div/ul/li[2]')
auckland_option.click()
positions_button = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[2]/div/div/div')
positions_button.click()
time.sleep(.5)
engineering_option = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[2]/div/div/ul/li[6]')
engineering_option.click()
while True:
    try:
        load_more = driver.find_element_by_xpath('//*[@id="JobsAjaxBtn"]')
        load_more.click()
        time.sleep(.5)
    except NoSuchElementException:
        break
jobs_of_interest = []
jobs = driver.find_elements_by_class_name('job')
for job in jobs:
    job_title = job.find_element_by_xpath('h3')
    name = job_title.text.upper()
    if 'PROPULSION' in name or 'JUNIOR' in name:
        link = job.get_attribute('href')
        print(name)
        print(link)
        jobs_of_interest.append((name, link))

# Writing and highlighting of new positions

# Take most recent Rocket Lab Positions data
previous_positions = pd.read_excel('data/' + sorted(os.listdir('data'))[-1])
previous_names = [name for name in previous_positions['Title']]


def highlight_cells(x):
    string = str(x).upper()
    if string not in previous_names and 'HTTPS' not in string:
        return "background-color: red"
    else:
        return "background-color: white"


path = f'data/{datetime.now().strftime("%Y-%m-%d--%H-%M-%S")}.xlsx'
writer = pd.ExcelWriter(path, engine='xlsxwriter')
new_positions = pd.DataFrame(jobs_of_interest, columns=['Title', 'URL'])
new_positions.style.applymap(highlight_cells).to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()
