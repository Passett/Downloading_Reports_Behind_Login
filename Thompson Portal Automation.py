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
import pandas as pd
import os
import csv
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

print ("Lee County report successfully downloaded")
logout() #Log out of portal

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
print("Sanibel report successfully downloaded")
driver.close()

print("Formatting data for reports")

#Assign file names to variables
fileNames=[]
for file in os.listdir(attachment_destination):
    fileNames.append(file)

Lee_PDF=attachment_destination+"/"+fileNames[0]
Myers_xlsx=attachment_destination+"/"+fileNames[1]
Sanibel_xlsx=attachment_destination+"/"+fileNames[2]

#Format the xlsx files
df_Myers_raw=pd.read_excel(Myers_xlsx, sheet_name=0)
df_Sanibel_raw=pd.read_excel(Sanibel_xlsx, sheet_name=0)

df_M1=df_Myers_raw.astype('str')
df_S1=df_Sanibel_raw.astype('str')

df_M2=df_M1.apply(lambda x: x.str.replace(',',' '))
df_S2=df_S1.apply(lambda x: x.str.replace(',',' '))

df_M3=df_M2.replace('nan',"")
df_S3=df_S2.replace('nan',"")

df_M3['Debris Class'] = df_M3['Debris Class'].replace([' and'], ' &', regex=True)
df_S3['Debris Class'] = df_S3['Debris Class'].replace([' and'], ' &', regex=True)

#Insert extra, blank columns
df_M3.insert(0, 'Project ID', 'FtMyersBeachFLIan')
df_M3.insert(1, 'Trailer Number', '')
df_M3.insert(2, 'Tare Tons', '')
df_M3.insert(3, 'Gross Tons', '')
df_M3.insert(4, 'Outbound Monitor', '')
df_M3.insert(5, 'Outbound LAT', '')
df_M3.insert(6, 'Outbound LONG', '')
df_M3.insert(7, 'Outbound Date', '')
df_M3.insert(8, 'Outbound Time', '')
df_M3.insert(9, 'Distance Direct', '')
df_M3.insert(10, 'Load Picture URL', '')
df_M3.insert(11, 'Inbound Picture URL', '')
df_M3.insert(12, 'Outbound Picture URL', '')
df_M3.insert(13, 'Zone Number', '')
df_M3.insert(14, 'Zone Name', '')
df_M3.insert(15, 'Street Number', '')
df_M3.insert(16, 'Street Name', '')
df_M3.insert(17, 'Road Owner', '')
df_M3.insert(18, 'Route ID', '')
df_M3.insert(19, 'Outbound Site Number', '')
df_M3.insert(20, 'Outbound Site Name', '')

df_S3.insert(0, 'Project ID', 'SanibelIslandFLIan')
df_S3.insert(1, 'Trailer Number', '')
df_S3.insert(2, 'Tare Tons', '')
df_S3.insert(3, 'Gross Tons', '')
df_S3.insert(4, 'Outbound Monitor', '')
df_S3.insert(5, 'Outbound LAT', '')
df_S3.insert(6, 'Outbound LONG', '')
df_S3.insert(7, 'Outbound Date', '')
df_S3.insert(8, 'Outbound Time', '')
df_S3.insert(9, 'Distance Direct', '')
df_S3.insert(10, 'Load Picture URL', '')
df_S3.insert(11, 'Inbound Picture URL', '')
df_S3.insert(12, 'Outbound Picture URL', '')
df_S3.insert(13, 'Zone Number', '')
df_S3.insert(14, 'Zone Name', '')
df_S3.insert(15, 'Street Number', '')
df_S3.insert(16, 'Street Name', '')
df_S3.insert(17, 'Road Owner', '')
df_S3.insert(18, 'Route ID', '')
df_S3.insert(19, 'Outbound Site Number', '')
df_S3.insert(20, 'Outbound Site Name', '')

#Rename column headers
dict={
    "Ticket No":"Ticket Number",
    "Debris Class":"Debris Type",
    "Capacity":"Truck Capacity",
    "Load %":"Percent Full",
    "Cubic Yards":"Pay Volume",
    "Weight":"Net Tons",
    "Truck No":"Truck Number",
    "Subcontractor":"Sub Contractor Code",
    "Load Latitude":"Load Lat",
    "Load Longitude":"Load Long",
    "Disposal Monitor":"Inbound Monitor",
    "Disposal Latitude":"Inbound LAT",
    "Disposal Longitude":"Inbound LONG",
    "Disposal Date":"Inbound Date",
    "Disposal Time":"Inbound Time",
    "Mileage":"Distance Haul Route",
    "Disposal Site":"Disposal Site ID"
    }

df_M3.rename(columns=dict, inplace=True)
df_S3.rename(columns=dict, inplace=True)

#Output only columns wanted, and in specified order
df_M4=df_M3[['Project ID', 'Ticket Number', 'Debris Type', 'Truck Capacity', 'Percent Full', 'Pay Volume', 'Gross Tons', 'Tare Tons', 'Net Tons', 'Truck Number', 'Trailer Number', 'Sub Contractor Code', 'Load Monitor', 'Load Lat', 'Load Long', 'Load Date', 'Load Time', 'Disposal Site ID', 'Inbound Monitor', 'Inbound LAT', 'Inbound LONG', 'Inbound Date', 'Inbound Time', 'Outbound Monitor', 'Outbound LAT', 'Outbound LONG', 'Outbound Date', 'Outbound Time', 'Distance Direct', 'Distance Haul Route', 'Load Picture URL', 'Inbound Picture URL', 'Outbound Picture URL', 'Zone Number', 'Zone Name', 'Street Number', 'Street Name', 'Road Owner', 'Route ID', 'Outbound Site Number', 'Outbound Site Name']]
df_S4=df_S3[['Project ID', 'Ticket Number', 'Debris Type', 'Truck Capacity', 'Percent Full', 'Pay Volume', 'Gross Tons', 'Tare Tons', 'Net Tons', 'Truck Number', 'Trailer Number', 'Sub Contractor Code', 'Load Monitor', 'Load Lat', 'Load Long', 'Load Date', 'Load Time', 'Disposal Site ID', 'Inbound Monitor', 'Inbound LAT', 'Inbound LONG', 'Inbound Date', 'Inbound Time', 'Outbound Monitor', 'Outbound LAT', 'Outbound LONG', 'Outbound Date', 'Outbound Time', 'Distance Direct', 'Distance Haul Route', 'Load Picture URL', 'Inbound Picture URL', 'Outbound Picture URL', 'Zone Number', 'Zone Name', 'Street Number', 'Street Name', 'Road Owner', 'Route ID', 'Outbound Site Number', 'Outbound Site Name']]

#Create DF for information they want above CSV, with the idea of adding this DF on top of our standard DF
yesterday_date=date.today()-timedelta(days = 1)
formatted_yesterday_date=yesterday_date.strftime("%#m/%#d/%Y")

M_data_above_header = {'Thompson Consulting':  ['FtMyersBeachFLIan', 'Debris Tickets', formatted_yesterday_date],
        '': ['', '1', '']
        }

S_data_above_header = {'Thompson Consulting':  ['SanibelIslandFLIan', 'Debris Tickets', formatted_yesterday_date],
        '': ['', '1', '']
        }

M_above_df = pd.DataFrame(M_data_above_header)
S_above_df = pd.DataFrame(S_data_above_header)

#Write CSVs that contains actual DFs, along with extra stuff requested above headers
#Ft Myers
with open((attachment_destination+'\\Ft Myers Beach LoadTickets_'+date.today().strftime("%m%d%Y")+'.csv'), 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(M_above_df.columns)
    for index,row in M_above_df.iterrows():
       writer.writerow(row)
    
    writer.writerow(df_M4.columns)
    for index1,row1 in df_M4.iterrows():
       writer.writerow(row1)
    
#Sanibel
with open((attachment_destination+'\\Sanibel LoadTickets_'+date.today().strftime("%m%d%Y")+'.csv'), 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(S_above_df.columns)
    for index,row in S_above_df.iterrows():
       writer.writerow(row)
    
    writer.writerow(df_S4.columns)
    for index1,row1 in df_S4.iterrows():
       writer.writerow(row1)

#Commented out, but below is how we would write to CSVs if we didn't need the extra stuff above the headers
# df_M4.to_csv(attachment_destination+'\\Ft Myers Beach LoadTickets_'+date.today().strftime("%m%d%Y")+'.csv', index=False)
# df_S4.to_csv(attachment_destination+'\\Sanibel LoadTickets_'+date.today().strftime("%m%d%Y")+'.csv', index=False)

Myers_csv=attachment_destination+'\\Ft Myers Beach LoadTickets_'+date.today().strftime("%m%d%Y")+'.csv'
Sanibel_csv=attachment_destination+'\\Sanibel LoadTickets_'+date.today().strftime("%m%d%Y")+'.csv'

print("prepping email")

#Open outlook and write email to Garrett and Buck, include subject, body, attachments
outlook = win32com.client.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'Sturgill.Simposon@metamodernsounds.com; Pooh.Shiesty@burrr.com'
mail.Subject = 'Daily Debris Reports'
mail.HTMLBody = '<h3>Greetings,<br><br>Please see the attached reports.<br><br>Sincerely,<br><br>Recovery</h3>'
mail.Body = "Greetings,\r\n\r\nPlease see the attached reports.\r\n\r\nSincerely,\r\n\r\nFDEM Recovery Bureau"
mail.Attachments.Add(Lee_PDF)
mail.Attachments.Add(Myers_csv)
mail.Attachments.Add(Sanibel_csv)
mail.CC = 'Shakey.Graves@ifnotforyou.com'
mail.Send()
print("Email sent, task complete!")
