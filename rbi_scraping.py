# -*- coding: utf-8 -*-
import os
import requests
import calendar


from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#path to chromedriver.exe
chrome_driver_path = 'C:/Users/ULTRAPC/Downloads/chromedriver_win32/chromedriver.exe'

#initialize ChromeDriver
service = Service(chrome_driver_path)
driver = webdriver.Chrome(service=service)

#create a directory to store the files
directory = 'RBI_Scraped_Files'
if not os.path.exists(directory):
    os.makedirs(directory)

#iterate through the months from April 2022 to March 2023
for year in range(2022, 2024):
    #open RBI website
    url = 'https://www.rbi.org.in/Scripts/ATMView.aspx'
    driver.get(url)

    #wait for the page to load
    wait = WebDriverWait(driver, 80)
    
    #find the year 
    year_btn = wait.until(EC.element_to_be_clickable((By.ID, f'btn{year}')))
    year_btn.click()

    #define the range for months for each year
    start_month = 4 if year == 2022 else 1
    end_month = 13 if year == 2022 else 4

    for month in range(start_month, end_month):
        #find the month in the accordion
        month_link_id = f'{year}{month}'
        month_link = wait.until(EC.presence_of_element_located((By.ID, month_link_id)))

        #stimulate the click action with javascript
        driver.execute_script("arguments[0].click();", month_link)

        #find the download link for the xlsx file of the corresponding month
        download_link = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'a[href$=".XLSX"]')))
        download_url = download_link.get_attribute('href')

        #download the excel file for the month
        response = requests.get(download_url)
        
        #get the month abbreviation for the file name
        month_abbr = calendar.month_abbr[month]

        #generate the file name with month abbreviation and year
        file_name = f'{month_abbr}_{year}.xlsx'
        file_path = os.path.join(directory, file_name)
        
        #save the downloaded file
        with open(file_path, 'wb') as file:
            file.write(response.content)

