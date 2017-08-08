import webbrowser
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
import openpyxl
from tqdm import tqdm
import time
'''
#--------------- Start Excel -----------------#
wb = openpyxl.Workbook()
wb = openpyxl.load_workbook(filename='policies.xlsm')
sheets = wb.sheetnames
ws = wb[sheets[0]]
#--------------- End Excel -------------------#
'''
pbar = tqdm(total=1)
#----------------- Read Excel Rows -----------------#    	
for i in range(3,4):
    try:
        index = str(i)
        driver = webdriver.Chrome()
        driver.get("https://tarif.allianz.ch/asu_cdn/apps/asu_momf-gui/#/de/")
        wait = WebDriverWait(driver,20)
        time.sleep(2)
        driver.find_element_by_xpath("/html/body/div[2]/div/div/div/div[1]/button").click()
        #element = wait.until(EC.element_to_be_clickable((By.ID, 'firstRegistration')))
        #driver.find_element_by_id("firstRegistration").send_keys("01.01.2017")
    except:
        pbar.close()
        pass
		#driver.close()
		#driver.quit()
    '''
	finally:
        pbar.update(1)
        driver.close()
        driver.quit()
    '''
pbar.close()