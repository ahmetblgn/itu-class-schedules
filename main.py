import time
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, UnexpectedAlertPresentException
from time import sleep
from openpyxl import Workbook
from selenium.common.exceptions import StaleElementReferenceException
import pandas as pd 
from selenium.webdriver.common.keys import Keys
import pyautogui as pt
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.alert import Alert

workbook = Workbook()
sheet = workbook.active

start_time = time.time()

class_list = ["AKM", "ALM", "ARB", "ARC", "ATA", "BBF", "BEB", "BED", "BEN", "BES", 
              "BIL", "BIO", "BLG", "BLS", "BUS", "CAB", "CEN", "CEV", "CHA", "CHE", 
              "CHZ", "CIE", "CIN", "CMP", "COM", "CVH", "DAN", "DEN", "DFH", "DGH", 
              "DNK", "DUI", "EAS", "ECN", "ECO", "EEE", "EEF", "EFN", "EHA", "EHB", 
              "EHN", "EKO", "ELE", "ELH", "ELK", "ELT", "END", "ENE", "ENG", "ENR", 
              "ENT", "ESL", "ESM", "ETK", "EUT", "FIZ", "FRA", "FZK", "GED", "GEM", 
              "GEO", "GID", "GLY", "GMI", "GMK", "GMZ", "GSB", "GSN", "GUV", "GVT", 
              "GVZ", "HSS", "HUK", "IAD", "ICM", "IEB", "ILT", "IML", "IND", "ING", 
              "INS", "ISE", "ISH", "ISL", "ISP", "ITA", "ITB", "JDF", "JEF", "JEO", 
              "JPN", "KIM", "KMM", "KMP", "KON", "LAT", "MAD", "MAK", "MAL", "MAR", 
              "MAT", "MCH", "MDN", "MEK", "MEN", "MET", "MIM", "MKN", "MMD", "MOD", 
              "MRE", "MRT", "MST", "MTH", "MTK", "MTM", "MTO", "MTR", "MUH", "MUK", 
              "MUT", "MUZ", "MYZ", "NAE", "NTH", "ODS", "PAZ", "PEM", "PET", "PHE", 
              "PHY", "PREP", "RES", "ROS", "RUS", "SBP", "SEC", "SED", "SEN", "SES", 
              "SGI", "SNT", "SPA", "STA", "STI", "TDW", "TEB", "TEK", "TEL", "TER", 
              "TES", "THO", "TRN", "TRS", "TRZ", "TUR", "UCK", "ULP", "UZB", "VBA", 
              "X100", "YTO", "YZV"]

class2_list = ["AKM", "ATA", "ING", "TUR", "UCK", "UZB"]

cService = webdriver.ChromeService(executable_path=r'C:\Users\Ahmet\Desktop\masaüstü\drivers\chromedriver.exe')
driver = webdriver.Chrome(service=cService)
result = []

index_for_course_code = 0
driver.get("https://obs.itu.edu.tr/public/DersProgram")
driver.maximize_window()

edu_xpath = "/html/body/div[1]/div[2]/div/div[1]/form/div/div[1]/span/span[1]/span/span[1]"
ders_xpath = "/html/body/div[1]/div[2]/div/div[1]/form/div/div[2]/span/span[1]/span/span[1]"
submit_xpath = "/html/body/div[1]/div[2]/div/div[1]/form/div/div[3]/button"
title_xpath = "/html/body/div[1]/div[2]/div/h1"

driver.find_element("xpath", title_xpath).click()

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

def safe_find_elements(xpath):
    try:
        elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, xpath))
        )
        return elements
    except Exception as e:
        print(f"Error: {e}")
        return []

while index_for_course_code <6:
    driver.get("https://obs.itu.edu.tr/public/DersProgram")
    driver.find_element("xpath", edu_xpath).click()
    sleep(1)
    arrow(1)
    pt.press("enter")
    sleep(3)

    try:
        driver.find_element("xpath", ders_xpath).click()
        pt.typewrite(class2_list[index_for_course_code])
        sleep(1)
        pt.press("enter")
        driver.find_element("xpath", submit_xpath).click()
        sleep(.3)

        # Check for alert before clicking on the page title
        try:
            alert = driver.switch_to.alert
            print(f"Alert detected: {alert.text}")
            sleep(.4)
            alert.accept()  # Click "OK"
            sleep(.4)
            index_for_course_code += 1
            continue  # Go to the next course code
        except:
            pass  # Continue if no alert is present

        driver.find_element("xpath", title_xpath).click()
        pt.press("enter")
        
        sleep(.5)  # Wait for the page to load

        # Get elements
        crns = safe_find_elements("//table/tbody/tr/td[1]")
        codes = safe_find_elements("//table/tbody/tr/td[2]")
        names = safe_find_elements("//table/tbody/tr/td[3]")
        teachingmethod = safe_find_elements("//table/tbody/tr/td[4]")
        instructors = safe_find_elements("//table/tbody/tr/td[5]")
        buildings = safe_find_elements("//table/tbody/tr/td[6]")
        days = safe_find_elements("//table/tbody/tr/td[7]")
        times = safe_find_elements("//table/tbody/tr/td[8]")
        rooms = safe_find_elements("//table/tbody/tr/td[9]")
        capacity = safe_find_elements("//table/tbody/tr/td[10]")
        enrolled = safe_find_elements("//table/tbody/tr/td[11]")
        restriction = safe_find_elements("//table/tbody/tr/td[13]")
        prerequisites = safe_find_elements("//table/tbody/tr/td[14]")
        creditrest = safe_find_elements("//table/tbody/tr/td[15]")

        for i in range(len(crns)):
            temp_data = {
                'Crn': crns[i].text,
                'Course Code': codes[i].text,
                'Course Name': names[i].text,
                'Teaching Method': teachingmethod[i].text,
                'Instructor': instructors[i].text,
                'Building': buildings[i].text,
                'Day': days[i].text,
                'Time': times[i].text,
                'Room': rooms[i].text,
                'Capacity': capacity[i].text,
                'Enrolled': enrolled[i].text,          
                'Restrictions': restriction[i].text,
                'Prerequisites': prerequisites[i].text,
                'Credit restrictions': creditrest[i].text,   
            }
            result.append(temp_data)

    except UnexpectedAlertPresentException:
        alert = driver.switch_to.alert
        print(f"Alert detected: {alert.text}")
        alert.accept()  # Close the alert
    except NoSuchElementException as e:
        print(f"Element not found: {e}")

    # Save to Excel every 10 iterations
    if (index_for_course_code + 1) % 10 == 0:
        df_data = pd.DataFrame(result)
        filename = f'C:/Users/Ahmet/Desktop/masaüstü/dersprogramı/bahar2425_{index_for_course_code + 1}.xlsx'
        df_data.to_excel(filename, index=False)
        print(f"Saved data to {filename}")

    index_for_course_code += 1  # Move to next course code

# Final save if there are remaining results
if result:
    df_data = pd.DataFrame(result)
    filename = f"C:/Users/Ahmet/Desktop/masaüstü/dersprogramı/class2_2425_final.xlsx"
    print(f"Saving final file: {filename}")
    df_data.to_excel(filename, index=False)
