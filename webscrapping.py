from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from openpyxl import Workbook
import smtplib
from email.message import EmailMessage
import os
from dotenv import load_dotenv
load_dotenv()

user_name = os.getenv('MAIL_ID')
pwd = os.getenv('PASS')

# Initialize the Chrome driver
serv_obj = Service("C:/chromedriver-win64/chromedriver.exe")  # Ensure correct path
driver = webdriver.Chrome(service=serv_obj)

# Open the Amazon website
driver.get("https://www.amazon.in/")

# Search for "samsung phone"
search_box = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
)
search_box.send_keys("samsung phone")
search_box.submit()

# Wait for the search results to load
WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CSS_SELECTOR, "span.a-size-medium.a-color-base.a-text-normal"))
)

# Collect phone names and prices
phone_names = driver.find_elements(By.XPATH, "//span[contains(text(),'Samsung Galaxy')]")
phone_prices = driver.find_elements(By.XPATH, "//span[contains(@class,'a-price-whole')]")

phone_name_list = []
phone_price_list = []

for phone_name in phone_names:
    phone_name_list.append(phone_name.text)

for phone_price in phone_prices:
    phone_price_list.append(phone_price.text)

# Combine names and prices into a list of tuples
phone_data = list(zip(phone_name_list, phone_price_list))

# Print the collected data
for data in phone_data:
    print(data)

# Close the browser
driver.quit()

# Create an Excel workbook and sheet
wb = Workbook()
wb.active.title = "Samsung Data"
sh1 = wb.active

# Append the header row
sh1.append(['Name', 'Price'])

# Append the phone data to the sheet
for item in phone_data:
    sh1.append(item)

# Save the workbook
wb.save("finalrecord.xlsx")


msg=EmailMessage()
msg['To'] = ['deepakaldo47@gmail.com','arvindarv01@gmail.com']
msg['From'] = 'aldoenterprise8@gmail.com'
msg['Subject'] = "training invitation"

with open('EmailTemplate.txt') as myfile:
    data=myfile.read()
    msg.set_content(data)

with open("finalrecord.xlsx","rb") as f: #read as binary
    file_data=f.read()
    file_name=f.name
    msg.add_attachment(file_data,maintype="application",subtype="xlsx",filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com',465) as server:
    server.login(user=user_name, password=pwd)
    server.send_message(msg)

print(" email sent!!")


