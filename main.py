import time
from datetime import datetime
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import os
# Only things you need to import is selenium for the webpage automation stuff
# (and download the chromedriver's appropriate version and put it somewhere)
# For the pandas/excel stuff you need pandas, openpyxl, and xlsxwriter


# Go to Career Positions page
driver = webdriver.Chrome('chromedriver/chromedriver.exe')
driver.get("https://www.rocketlabusa.com/careers/positions/")

# Select relevant filters
location_dropdown = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[1]/div/div/div')
location_dropdown.click()
# Without this delay it sometimes tries to find the element before it has loaded and returns NoSuchElementException
time.sleep(.5)
auckland_option = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[1]/div/div/ul/li[2]')
auckland_option.click()
positions_button = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[2]/div/div/div')
positions_button.click()
time.sleep(.5)
engineering_option = driver.find_element_by_xpath('//*[@id="Jobs"]/div/div[1]/form/div[2]/div/div/ul/li[6]')
engineering_option.click()
# Press "Load More" button as long as its an option
while True:
    try:
        load_more = driver.find_element_by_xpath('//*[@id="JobsAjaxBtn"]')
        load_more.click()
        time.sleep(.5)
    except NoSuchElementException:
        break

# Get job urls and titles and save in list
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


# Take most recent Rocket Lab Positions data
previous_positions = pd.read_excel('data/' + sorted(os.listdir('data'))[-1])
previous_names = [name for name in previous_positions['Title']]


# Writing and highlighting of new positions
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
