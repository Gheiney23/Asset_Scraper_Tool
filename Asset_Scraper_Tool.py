import os
import pandas as pd
import shutil
import time
from openpyxl import load_workbook
from selenium import webdriver as wb
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
import selenium.webdriver.support.ui as ui
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from urllib.request import urlretrieve

# List of skus
sku_list = [
'Skus_List'
]

# Urls for the skus on the manufacturer website
url_path = [
'Sku_URL_List'
]

# Creating a folder to store assets on the desktop
path = 'Folder_path'
os.mkdir(path)

# Setting up the webdriver for Selenium
options = wb.ChromeOptions()
options.add_argument('--start-maximized')
options.add_experimental_option('excludeSwitches', ['enable-logging'])
driver = wb.Chrome(options=options)

# Creating a function to scrap the image assets from the webpage and group them by sku
src_list = []
skus = []
file_names = []
def get_assets(url, sku):
    driver.get(url)
    time.sleep(1)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="lk-product-media-carousel"]')))
    driver.find_element_by_xpath('//*[@id="carouselImages"]/div/div[1]/a/img').click()
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="imageModalImage"]')))
    driver.find_element_by_xpath('//*[@id="imageModal"]/div/div/div[3]/a[1]/img').click()
    
    driver.find_element_by_xpath('//*[@id="imageModal"]/div/div/div[3]/a[1]/img').click()
    img_1 = driver.find_element_by_xpath('//*[@id="imageModalImage"]')
    src_1 = img_1.get_attribute('src')
    skus.append(sku)
    src_list.append(src_1)
    file_name = sku + "_" + str(len(src_list))
    urlretrieve(src_1, 'Folder_path\\{}.jpg'.format(file_name))
    file_names.append(file_name)

    driver.find_element_by_xpath('//*[@id="imageModal"]/div/div/div[3]/a[2]/img').click()
    img_2 = driver.find_element_by_xpath('//*[@id="imageModalImage"]')
    src_2 = img_2.get_attribute('src')
    src_list.append(src_2)
    skus.append(sku)
    file_name = sku + "_" + str(len(src_list))
    urlretrieve(src_2, 'Folder_path\\{}.jpg'.format(file_name))
    file_names.append(file_name)
    
    driver.find_element_by_xpath('//*[@id="imageModal"]/div/div/div[3]/a[3]/img').click()
    img_3 = driver.find_element_by_xpath('//*[@id="imageModalImage"]')
    src_3 = img_3.get_attribute('src')
    src_list.append(src_3)
    skus.append(sku)
    file_name = sku + "_" + str(len(src_list))
    urlretrieve(src_3, 'Folder_path\\{}.jpg'.format(file_name))
    file_names.append(file_name)
    
    driver.find_element_by_xpath('//*[@id="imageModal"]/div/div/div[3]/a[4]/img').click()
    img_4 = driver.find_element_by_xpath('//*[@id="imageModalImage"]')
    src_4 = img_4.get_attribute('src')
    src_list.append(src_4)
    skus.append(sku)
    file_name = sku + "_" + str(len(src_list))
    urlretrieve(src_4, 'Folder_path\\{}.jpg'.format(file_name))
    file_names.append(file_name)

get_assets(url_path[0], sku_list[0])
get_assets(url_path[1], sku_list[1])
get_assets(url_path[2], sku_list[2])
get_assets(url_path[3], sku_list[3])
get_assets(url_path[4], sku_list[4])
get_assets(url_path[5], sku_list[5])
get_assets(url_path[6], sku_list[6])
get_assets(url_path[7], sku_list[7])
get_assets(url_path[8], sku_list[8])
get_assets(url_path[9], sku_list[9])
get_assets(url_path[10], sku_list[10])

#  Resizes the asset with new url parameters
new_src_list = []
for src in src_list:
    new_src = src.replace('1000x1000boundedresize', '2000x2000boundedresize')
    new_src_list.append(new_src)

# Creating a dictionary from the lists and loading them into a DataFrame
d = {'Sku': skus, 'File_name': file_names, 'Img_url': new_src_list}    
src_df = pd.DataFrame(d)


driver.quit()

# Creating a zipped folder from the original asset folder on the desktop
shutil.make_archive('Folder_path')

# Loading the DataFrame into Excel as a worksheet
path = 'Excel_file'
excel_wb = load_workbook(path)
with pd.ExcelWriter(path) as writer:
    writer.book = excel_wb
    src_df.to_excel(writer, sheet_name='Img_url_data')