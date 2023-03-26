from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pandas as pd
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException

PERMITS_URL = "https://data.opendatasoft.com/explore/dataset/issued-building-permits%40vancouver/export/"
STATUS_URL = "https://plposweb.vancouver.ca/Public/Default.aspx?PossePresentation=PermitSearchByNumber"
DOWNLOAD_PATH = r"C:\Users\durga\Downloads\issued-building-permits@vancouver.xlsx"

PATH = "C:\Program Files (x86)\chromedriver.exe"
driver = webdriver.Chrome(PATH)

def getPermitsInExcel():
   driver.get(PERMITS_URL)
   search = driver.find_element(By.XPATH, "//a[@aria-label='Dataset export (Excel)']")
   search.send_keys(Keys.RETURN)
   time.sleep(15)

def fetchStatusOfPermit(PermitNo):
   driver.get(STATUS_URL)
   search = driver.find_element(By.ID, 'ExternalFileNum_1003032_S0')
   search.send_keys(PermitNo)
   search.send_keys(Keys.RETURN)
   time.sleep(2)
   try:
      main = driver.find_element(By.CSS_SELECTOR, "span[id*='StatusDescription']")
      res = str(main.text)
      # go back
      back = driver.find_element(By.CSS_SELECTOR, "a[id*='Search']").click()
   except NoSuchElementException:
      res = 'Not found'
   
   return res

def readAndWriteExcel():
   df = pd.read_excel(DOWNLOAD_PATH)
   lst_permit_numbers = list(df['PermitNumber'])
   # lst_permit_numbers = ['BP-2021-04891']
   # df = pd.DataFrame(lst_permit_numbers, columns=['Numbers'])
   lst_permit_status = []
   try: 
      for i in lst_permit_numbers:
         lst_permit_status.append(fetchStatusOfPermit(i))
   except NoSuchElementException:
      pass
   df['Status'] = lst_permit_status
   df.to_excel('withStatus.xlsx', sheet_name='Sheet1')

   print(lst_permit_status)


def keepOpen():
   while True:
    pass

readAndWriteExcel()
keepOpen()



