from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
import time
from time import sleep
from openpyxl import Workbook
import pandas as pd 
from selenium.webdriver.common.keys import Keys
import pyautogui as pt

cService = webdriver.ChromeService(executable_path=r'C:\Users\Ahmet\Desktop\masaüstü\drivers\chromedriver.exe')
driver = webdriver.Chrome(service=cService)
driver.get("https://obs.itu.edu.tr/public/DersProgram")

edu_xpath = "/html/body/div[1]/div[2]/div/div[1]/form/div/div[1]/span/span[1]/span/span[1]"
ders_xpath = "/html/body/div[1]/div[2]/div/div[1]/form/div/div[2]/span/span[1]/span/span[1]"
submit_xpath = "/html/body/div[1]/div[2]/div/div[1]/form/div/div[3]/button"
title_xpath = "/html/body/div[1]/div[2]/div/h1"
def arrow(quantity):
    for i in range(quantity):
        pt.press("down")

def uparrow(quantity):
    for i in range(quantity):
        pt.press("up")

driver.find_element("xpath", edu_xpath).click()
sleep(1)
arrow(1)
pt.press("enter")

sleep(3)

# Dropdown'u aç
#dropdown = driver.find_element(By.CLASS_NAME, "select2-selection--single")
dropdown = driver.find_element("xpath", "/html/body/div[1]/div[2]/div/div[1]/form/div/div[2]/span/span[1]/span/span[1]")

dropdown.click()
time.sleep(1)  # Açılması için bekle

# Tüm seçenekleri bul
options = driver.find_elements(By.CSS_SELECTOR, ".select2-results__option")

# Seçenekleri yazdır
for option in options:
    print(option.text)

driver.quit()
