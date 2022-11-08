#This script was written by Richard Passett and downloads the files for Johnny from the Thompson Portal and sends them to Buck and Garrett. 
#It currently grabs 3 files using 3 different logins.

#Import dependencies
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
import shutil
import keyring
from datetime import date
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

#Thompson Report Listings used for script
Accounts_Listing="https://floridapa.org/app/#account/accountlist?o=grantname+asc%2Capplicantname+asc"
Appeals_Listing="https://floridapa.org/app/#project/projectappeallist?"
large_project_closeout_listing="https://floridapa.org/app/#project/projectcloseoutlist?filters=%7B%22Program%22%3A%221%22%2C%22Step%22%3A%2226%2C27%2C28%2C570%2C29%2C482%2C485%2C157%2C571%2C183%2C446%2C572%2C159%22%7D&pp=25&o=laststepchangedays+asc"

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
#options.add_argument("--headless")
#options.add_argument("--disable-software-rasterizer")
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
    counter=1
    while len(os.listdir(holding_dir))==0: 
        time.sleep(10)
        counter+=counter
        if counter==8:
            sys.exit("Today's data hasn't been uploaded yet. Please try again later.")
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

#Prep Folder
cleanFolder(attachment_destination)

#Download 1st report
#Login and go to report page
login(Johnny_username_1, Johnny_password)
time.sleep(5)
part1="https://portal.thompsoncs.net/documents/reports/E9A1B06B-D7AE-4379-B40F-91B375E5E146/"
part2=date.today().strftime("%#m.%#d.%y")
part3="_Ft Myers Beach Daily Report.pdf"
Lee_Listing=(part1+part2+part3)
driver.get(Lee_Listing)
time.sleep(10)
move(attachment_destination)
print("1st file downloaded")

#Log out of portal
logout_button=driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[63]/div/table/tbody/tr/td[2]')
logout_button.click()

#Download 2nd Report
#Login and download report
login(Johnny_username_3, Johnny_password)
time.sleep(10)
part1="https://portal.thompsoncs.net/documents/reports/9E9BF0D9-A98F-4F4D-B160-EB16D1062528/"
part2=date.today().strftime("%#m.%#d.%y")
part3="_Sanibel Daily Report.pdf"
Lee_Listing=(part1+part2+part3)
driver.get(Lee_Listing)
time.sleep(10)
move(attachment_destination)
print("2nd file downloaded")

#Log out of portal
logout_button=driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div[3]/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr/td[63]/div/table/tbody/tr/td[2]')
logout_button.click()

#Download 3rd Report
#Login and go to report page
login(Johnny_username_2, Johnny_password)
driver.get("https://portal.thompsoncs.net/reports.aspx")
time.sleep(5)
part1="https://portal.thompsoncs.net/documents/reports/1466F032-87E2-4929-9B4E-63E6C7861DC1/"
part2=date.today().strftime("%#m.%#d.%y")
part3="_Lee County Daily Report.pdf"
Lee_Listing=(part1+part2+part3)
driver.get(Lee_Listing)
move(attachment_destination)
driver.close()
print("3rd file downloaded")

#Assign file names to variables
fileNames=[]
for file in os.listdir(attachment_destination):
    fileNames.append(file)

attachment1=attachment_destination+"/"+fileNames[0]
attachment2=attachment_destination+"/"+fileNames[1]
attachment3=attachment_destination+"/"+fileNames[2]

#Open outlook and write email to Garrett and Buck, include subject, body, attachments
print("prepping email")
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'richard.passett@em.myflorida.com'
mail.Subject = 'Daily Debris Reports'
mail.HTMLBody = '<h3>Greetings,<br><br>Please see the attached reports.<br><br>Sincerely,<br><br>Recovery</h3>'
mail.Body = "Greetings,\r\n\r\nPlease see the attached reports.\r\n\r\nSincerely,\r\n\r\nFDEM Recovery Bureau"
mail.Attachments.Add(attachment1)
mail.Attachments.Add(attachment2)
mail.Attachments.Add(attachment3)
#mail.CC = 'somebody@company.com'
mail.Send()
print("Email sent, task complete!")
