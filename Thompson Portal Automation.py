#This script was written by Richard Passett and downloads the files for Johnny from the Thompson Portal and sends them to Buck and Garrett. 
#It currently grabs 3 files using 3 different logins.

#Import dependencies
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import shutil
import keyring
from datetime import date, datetime, timedelta
import win32com.client
import os
import sys

#Password variables for Thompson Portal
Johnny_username_1=keyring.get_password("TP", "Username_1")
Johnny_username_2=keyring.get_password("TP", "Username_2")
Johnny_username_3=keyring.get_password("TP", "Username_3")
Johnny_password=keyring.get_password("TP", "Password")

#Directories used for script
holding_dir=r'J:\Admin & Plans Unit\Recovery Systems\3. Projects\Johnny_Automation\Holding_Folder'
attachment_destination=r'J:\Admin & Plans Unit\Recovery Systems\3. Projects\Johnny_Automation\Attachments'

#Variables for choosing dropdown options
sanibel="san"
Myers="ft"
lee='le'

#Use webdriver for Chrome, set where you want the CSVs to download to, add other options/preferences as desired, point to where you have the driver downloaded, and set the driver to a variable.
#If you want to see what is happening in the browser, comment out the headless and disable-software-rasterizer options
options=webdriver.ChromeOptions()
prefs={
    'download.prompt_for_download': False,
    "download.default_directory" : r'J:\Admin & Plans Unit\Recovery Systems\3. Projects\Johnny_Automation\Holding_Folder',
    'download.directory_upgrade': True,
    'plugins.always_open_pdf_externally': True
    }
options.add_experimental_option("prefs",prefs) 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--headless")
options.add_argument("--disable-software-rasterizer")
options.add_argument("--start-maximized")
driver_service=Service(r"C:\Users\richardp\Desktop\chromedriver\chromedriver.exe")
driver=webdriver.Chrome(service=driver_service, options=options)
wait=WebDriverWait(driver, 120)

#Functions used for automation
#Log into portal
def login(username, password):
    driver.get("https://portal.thompsoncs.net/login.aspx")
    time.sleep(8)
    username_field=driver.find_element(By.NAME,"u1")
    password_field=driver.find_element(By.NAME,"u2")
    signIn_button=driver.find_element(By.NAME,"btnAuthenticate")
    username_field.clear()
    password_field.clear()
    username_field.send_keys(username)
    password_field.send_keys(password)
    signIn_button.click()
    time.sleep(5)

def cleanFolder(destination):
    for file in os.scandir(destination):
        os.remove(file.path)    

def move(destination):
    while len(os.listdir(holding_dir))==0: 
        time.sleep(10)
    for item in os.listdir(holding_dir):
        file_name=holding_dir+"/"+item
        if item.endswith(".tmp"):
            time.sleep(10)
            move(destination)
        elif item.endswith("crdownload"):
            time.sleep(10)
            move(destination)
        else:
            shutil.copy2(file_name, destination) #Copy csv to JDrive
            os.remove(file_name) #Delete original file
    time.sleep(5)

def logout():
    logout_button=driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[63]/div/table/tbody/tr/td[2]')
    logout_button.click()

#Prep Folder
cleanFolder(attachment_destination)

print("Greetings, we are pulling your reports for you now")

#Download 1st Report
#Login and download Lee report
login(Johnny_username_2, Johnny_password)
driver.get("https://portal.thompsoncs.net/reports.aspx")
time.sleep(5)

#Choose Lee County
dropdown_button=driver.find_element(By.NAME,'ctl00$MainContent$ddlFilterClient')
dropdown_button.send_keys(lee)
time.sleep(1)
dropdown_button.send_keys(Keys.ENTER)
time.sleep(8)

#Click filter button
filter_button=driver.find_element(By.NAME,'ctl00$MainContent$btnFilter')
try:
    filter_button.click()
except:
    driver.find_element(By.NAME,'ctl00$MainContent$btnFilter').click()
time.sleep(8)

#Grab today's file and move it to correct folder
#Find latest "Document Date" from site and make sure that it equals yesterday's date.
#If so, download file. If not, stop script and let user know today's file isn't ready yet 
goodies=date.today().strftime("%#m.%#d.%y")
yesterday_prep=date.today()-timedelta(days = 1)
alternate_goodies=yesterday_prep.strftime("%#m.%#d.%y")
latest_in_system=driver.find_element(By.XPATH, '/html/body/form/div[4]/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table[2]/tbody/tr/td/div/table[2]/tbody/tr[4]/td[3]').text
start_formatting=latest_in_system.strip()
check=datetime.strptime(start_formatting, "%m/%d/%Y")
formatted_check=check.strftime("%#m.%#d.%y")

if alternate_goodies==formatted_check:
    try:
        driver.find_element(By.XPATH, '//a[contains(@href, "%s")]' % goodies).click()
        move(attachment_destination)
    except:
        try:
            driver.find_element(By.XPATH, '//a[contains(@href, "%s")]' % alternate_goodies).click()
            move(attachment_destination)
        except:
            sys.exit("We ran into an issue locating the pdf report for Lee County due to a format change on Thompson's side. Please investigate or contact Richard.Passett@em.myflorida.com")
else:
    sys.exit("Today's data has not yet been updated")

driver.close()
print ("Lee County report successfully downloaded")

#Download 2nd report
#Login and download Ft. Myers Report
login(Johnny_username_1, Johnny_password)
driver.get("https://portal.thompsoncs.net/tickets.aspx")
time.sleep(5)

#Choose ft myers
dropdown_button=driver.find_element(By.NAME,'ctl00$MainContent$ddlFilterLoadClient')
dropdown_button.send_keys(Myers)
time.sleep(1)
dropdown_button.send_keys(Keys.ENTER)
time.sleep(8)

#Click filter button
filter_button=driver.find_element(By.NAME,'ctl00$MainContent$btnLoadFilter')
try:
    filter_button.click()
except:
    driver.find_element(By.NAME,'ctl00$MainContent$btnLoadFilter').click()
time.sleep(8)

#Click "Export to Excel" button and move download to correct folder
download_button=driver.find_element(By.NAME,'ctl00$MainContent$btnLoadExcel')
try:
    download_button.click()
except:
    driver.find_element(By.NAME,'ctl00$MainContent$btnLoadExcel').click()
move(attachment_destination)
logout() #Log out of portal
print("Ft. Myers report successfully downloaded")

#Download 3rd Report
#Login and download Sanibel report
login(Johnny_username_3, Johnny_password)
driver.get("https://portal.thompsoncs.net/tickets.aspx")
time.sleep(8)

#Choose sanibel
dropdown_button=driver.find_element(By.NAME,'ctl00$MainContent$ddlFilterLoadClient')
dropdown_button.send_keys(sanibel)
dropdown_button.send_keys(Keys.ENTER)
time.sleep(2)

#Click filter button
filter_button=driver.find_element(By.NAME,'ctl00$MainContent$btnLoadFilter').click()
try:
    filter_button.click()
except:
    driver.find_element(By.NAME,'ctl00$MainContent$btnLoadFilter').click()
time.sleep(10)

#Click "Export to Excel" button and move download to correct folder
download_button=driver.find_element(By.NAME,'ctl00$MainContent$btnLoadExcel')
try:
    download_button.click()
except:
    driver.find_element(By.NAME,'ctl00$MainContent$btnLoadExcel').click()
move(attachment_destination)
logout() #Log out of portal
print("Sanibel report successfully downloaded")

print("prepping email")

#Assign file names to variables
fileNames=[]
for file in os.listdir(attachment_destination):
    fileNames.append(file)

attachment1=attachment_destination+"/"+fileNames[0]
attachment2=attachment_destination+"/"+fileNames[1]
attachment3=attachment_destination+"/"+fileNames[2]

#Open outlook and write email to Garrett and Buck, include subject, body, attachments
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gsauls@debristech.com; bdickinson@debristech.com'
mail.Subject = 'Daily Debris Reports'
mail.HTMLBody = '<h3>Greetings,<br><br>Please see the attached reports.<br><br>Sincerely,<br><br>Recovery</h3>'
mail.Body = "Greetings,\r\n\r\nPlease see the attached reports.\r\n\r\nSincerely,\r\n\r\nFDEM Recovery Bureau"
mail.Attachments.Add(attachment1)
mail.Attachments.Add(attachment2)
mail.Attachments.Add(attachment3)
mail.CC = 'jonathan.blocker@em.myflorida.com'
mail.Send()
print("Email sent, task complete!")
