#To Do
#Add general description comment at top of code, clean up imports, and write better comments throughout
#Create a FLPA login that is used only for reports and change the login info in this code to the new login

#Import dependencies
import time
from datetime import date
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from zipfile import ZipFile
import os
import shutil
import keyring
import win32com.client

#Password variables for FLPA and Grants Portal
my_username=keyring.get_password("FLPA_GP", "username")
FLPA_password=keyring.get_password("FLPA", "Passett")

#Use webdriver for Chrome and set where you want the csv to download to
options=webdriver.ChromeOptions()
prefs={"download.default_directory" : r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Holding Folder'}
options.add_experimental_option("prefs",prefs) 
options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument("--headless")
options.add_argument("--disable-software-rasterizer")
driver_service=Service(r"C:\Users\richardp\Desktop\chromedriver\chromedriver.exe")
driver=webdriver.Chrome(service=driver_service, options=options)
wait=WebDriverWait(driver, 120)

#Function that downloads CSV files. 
#The process is the same with the same locations for all small reports, which is why we can build a reusable function for this.
def download_report():
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR,"div.toExcel.inner")))
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"div.toExcel.inner")))
    Excel_button=driver.find_element(By.CSS_SELECTOR,"div.toExcel.inner")
    driver.execute_script("arguments[0].click();", Excel_button)
    wait.until(EC.element_to_be_clickable((By.ID,'excelexportcolumns2')))
    Custom_button=driver.find_element(By.ID,'excelexportcolumns2')
    driver.execute_script("arguments[0].click();", Custom_button)
    wait.until(EC.element_to_be_clickable((By.CLASS_NAME,'selectAll')))
    selectAll_button=driver.find_element(By.CLASS_NAME,'selectAll')
    driver.execute_script("arguments[0].click();", selectAll_button)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR,"input.close.main")))
    export_button=driver.find_element(By.CSS_SELECTOR,"input.close.main")
    driver.execute_script("arguments[0].click();", export_button)
    time.sleep(2)
    try:
        driver.find_elements(By.CSS_SELECTOR,"input.close.main")[-1].click()
    except IndexError:
        pass

#function to move csv to desired destination. Waits for file to exist, empties destination folder before moving new file, and accounts for whether or not csv is in a zip file.
def move(destination):
    while len(os.listdir(dir_name))==0: 
        time.sleep(10)
    for file in os.scandir(destination):
        os.remove(file.path)
    for item in os.listdir(dir_name):
        file_name=dir_name+"/"+item
        if item.endswith(".zip"):
            zip_ref = ZipFile(file_name) # create zipfile object
            zip_ref.extractall(destination) # extract file to dir
            zip_ref.close() # close file
            os.remove(file_name) #Delete original file
        elif item.endswith("crdownload"):
            time.sleep(10)
            move(destination)
        else:
            shutil.copy2(file_name, destination) #Copy csv to JDrive
            os.remove(file_name) #Delete original file
    time.sleep(5)

#Function to download FLPA CSVs. Accepts three arguments; driver.get location, destination path, and sleep time 
def export(listing, destination):
    driver.get(listing)
    time.sleep(20)
    download_report()
    move(destination)

#Provide a message to the person running this script
print("Greetings, we are pulling your Payables Report data for you now.\nThis will take about 5 minutes and we will let you know as soon as this task is complete.")

#Open FLPA
driver.get("https://floridapa.org/")
time.sleep(5)

#Login to FLPA
username_field=driver.find_element(By.NAME,"Username")
password_field=driver.find_element(By.NAME,"Password")
signIn_button=driver.find_element(By.NAME,"Submit")
username_field.clear()
password_field.clear()
username_field.send_keys(my_username)
password_field.send_keys(FLPA_password)
signIn_button.click()
time.sleep(5)

#Go to Payables Report Data. We cut the html into 7 parts so that we can replace the end-date on the filter with the date that the program runs
p1="https://floridapa.org/app/#payment/payablelist?filters=%7B%22Step%22%3A%22575%2C8%2C211%2C458%2C460%2C459%2C9%22%2C%22SubmittedDate%22%3A%22Jan+8%2C+2019----"
p2=date.today().strftime("%b")
p3="+"
p4=date.today().strftime("%d")
p5="%2C+"
p6=date.today().strftime("%Y")
p7="%22%7D&o=laststepchangedays+asc&p=1&pp=50&s="
payables_filters=(p1+p2+p3+p4+p5+p6+p7)

Payables_Destination=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\FLPA Payables Export'
dir_name=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\4. Data Files\Holding Folder'

export(payables_filters, Payables_Destination)

driver.close()

# Open the report template, refresh the data sources, delete queries from workbook, save as new name in correct location

today=date.today().strftime("%Y%m%d")
filename=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\2. Weekly\Payables Report\Payables_Report_Template.xlsx'
newfile=r'J:\Admin & Plans Unit\Recovery Systems\2. Reports\2. Weekly\Payables Report\Archive\2022\\'+"Payables_Report_"+today+".xlsx"

xl = win32com.client.DispatchEx("Excel.Application")
wb = xl.Workbooks.Open(filename)
xl.Visible = True
wb.RefreshAll()
xl.CalculateUntilAsyncQueriesDone()
time.sleep(15)
for c in wb.Connections:
    c.Delete()
for q in wb.Queries:
     q.Delete()
wb.SaveAs(newfile)
wb.Close(True)
xl.Quit()

#Start the email for the report

# outlook = win32com.client.gencache.EnsureDispatch('Outlook.Application')
# new_mail = outlook.CreateItem(0)

# # Label the subject
# new_mail.Subject = "Payables Report {:%m-%d-%y}".format(date.today())

# # Add the to and cc list
# to_email="""
# Melissa Shirah <Melissa.Shirah@em.myflorida.com>; Douglas Roberts <Douglas.Roberts@em.myflorida.com>; Ben Fairbrother <Ben.Fairbrother@em.myflorida.com>; Ronald Baker <Ronald.Baker@em.myflorida.com>; Jennifer Stallings <Jennifer.Stallings@em.myflorida.com>
# """
# cc_email="""
# Jonathan Blocker <Jonathan.Blocker@em.myflorida.com>; Christopher Sabol <Christopher.Sabol@em.myflorida.com>
# """
# new_mail.To=to_email
# new_mail.CC=cc_email

# # Attach the file (Needs to be a string, not a path)
# new_mail.Attachments.Add(Source=str(newfile))

# message="Greetings,\nAttached is the most recent payables report. Please let me know if you have any questions or concerns."
# new_mail.Display(True)

print("Task Complete")