from selenium import webdriver
import time
import os
import datetime
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import ElementNotInteractableException
import pandas as pd
options = webdriver.ChromeOptions()
prefs = {"download.default_directory" : "C:\_temp"}
options.add_experimental_option("prefs",prefs)
options.add_argument("--start-maximized")
service = Service(executable_path='C:\Chromedriver\chromedriver')
driver = webdriver.Chrome(service=service, options=options)
MSTR = driver.get("https://mstrprod.mhf.mhc/MicroStrategy/servlet/mstrWeb")
driver.find_element(By.XPATH, '//*[@id="projects_ProjectsStyle"]/table/tbody/tr[2]/td[1]/div/table/tbody/tr/td[2]/a').click()
driver.find_element(By.XPATH, '/html/body/div[4]/table/tbody/tr[2]/td[2]/div[2]/div[1]/a[1]/div[1]').click()
driver.find_element(By.XPATH, "/html/body/div[4]/table/tbody/tr[2]/td[2]/div[2]/div[1]/div/table/tbody/tr[1]/td[1]/div/table/tbody/tr/td[2]/a").click()
driver.find_element(By.XPATH, "/html/body/div[4]/table/tbody/tr[2]/td[2]/div[2]/div[1]/div/table/tbody/tr[3]/td/div/table/tbody/tr/td[2]/a").click()
driver.find_element(By.XPATH, "/html/body/table/tbody/tr[2]/td/div[5]/div/div/div/div[2]/div/span[1]/span/span/table/tbody/tr/td[40]/div/table/tbody/tr/td[3]/div/div").click()
driver.find_element(By.XPATH, "/html/body/table/tbody/tr[2]/td/div[5]/div/div/div/div[2]/div/span[1]/span/span/table/tbody/tr/td[40]/table/tbody/tr/td/span/div[3]/div[1]/div[1]/div").click()
driver.find_element(By.XPATH, "/html/body/div[4]/table/tbody/tr[2]/td[2]/div[2]/div[3]/div[1]/table/tbody/tr[2]/td/table/tbody/tr[1]/td/table/tbody/tr[1]/td[2]/input[1]").click()
driver.find_element(By.XPATH, "/html/body/div[4]/table/tbody/tr[2]/td[2]/div[2]/div[3]/div[1]/table/tbody/tr[2]/td/table/tbody/tr[2]/td/input[1]").click()
file = pd.read_excel(r"C:\_temp\Pipeline Report - Default Filters.xlsx")
file.to_excel(r'Y:\Data Operations\Data Operations RMBS_ABS_CMBS\Servicer Outreach\4. Pipeline reports\Pipeline Report.xlsx')