from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl



driver=webdriver.Chrome('web driver for chrome is needed put the PATH here.exe')
driver.get('put the website here.com')
driver.maximize_window()

## CSS selector of each cell
employees=[
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T2"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T3"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D6"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D7"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D8"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T10"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T11"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T13"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T15"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D16"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T18"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T19"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D21"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D34"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D37"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D39"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D40"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D41"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D43"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D44"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_T45"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D46"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D47"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D48"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D49"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D51"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D52"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_D54"]',
           '//*[@id="ctl00_m_g_c5baf3ee_3f9c_4a5d_b8fb_4be97d49decf_FormControl0_V1_I1_RTC29_RTI5_RT1_newRichText"]/p']
        
          
# fill in each cell with the following info from a excel file.      
excel=[]     
for i in range(1,31): 
    wb=openpyxl.load_workbook('Book1.xlsx')
    sheet= wb.get_sheet_by_name('Sheet1')
    excel.append(str(sheet.cell(row=1,column=i).value)) 
for key,value in zip(employees,excel):
      element =driver.find_element_by_xpath(key)
      actions=ActionChains(driver)
      actions.move_to_element(element)
      actions.click()
      actions.send_keys(value)
      actions.perform()
      time.sleep(.2)
          
                  
