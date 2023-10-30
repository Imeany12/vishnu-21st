# Import required libraries
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager 

# Read data from the input Excel file
df = pd.read_excel("rawData.xlsx file", sheet_name="sheet name")

# Insert a new column named "school location" at a specific position
df.insert(df.columns.get_loc("โรงเรียนที่สำเร็จการศึกษา ม.6 (Graduated high school)"), "school location", "")

# Set options for the Chrome webdriver
options = Options()
service = Service(ChromeDriverManager().install())
options.add_experimental_option("detach", True)  # Keep the browser open after the script execution
driver = webdriver.Chrome(service=service, options=options)

# Navigate to the web page to be scraped
driver.get("http://reg4.sut.ac.th/registrar/ces_SearchSchool.asp")

# Define a function to scrape data based on the school name
def scrape(schoolName):
    # Find the input field and submit button on the web page
    inputSch = driver.find_element(By.NAME, "nameschool")
    summit = driver.find_element(By.XPATH, "/html/body/table[1]/tbody/tr/th/table/tbody/tr[3]/td/form/input[2]")

    # Enter the school name into the input field and click the submit button
    inputSch.send_keys(schoolName)
    summit.click()

    # Find and extract the province information from the web page
    provice = driver.find_element(By.XPATH, "/html/body/table[2]/tbody/tr[2]/td[2]/font")
    return provice.text

# Iterate over each cell in the DataFrame and perform the scraping operations
for cell in df.index:
    try:
        # Extract the school name from the DataFrame and scrape the province information
        name = str(df.loc[cell, "โรงเรียนที่สำเร็จการศึกษา ม.6 (Graduated high school)"]).strip()
        province = scrape(name)
        df.loc[cell, "school location"] = province  # Assign the scraped province information to the DataFrame
    except Exception as e:
        # If an exception occurs, print the cell index and school name for debugging purposes
        print(f"{cell} {name}")

# Close the webdriver
driver.close()

# Write the resulting DataFrame to a new Excel file
df.to_excel("resultSchool.xlsx", index=False)
