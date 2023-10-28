# run pip install -U selenium first to update selenium
# Import necessary libraries
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options

# Read data from the input Excel file and perform data cleaning
data = pd.read_excel("rawData.xlsx")
data.dropna(inplace=True)  # Remove any rows with missing values
data.drop_duplicates(subset=["ชื่อ-นามสกุล ภาษาไทย (Full name in Thai)"], keep="last", inplace=True)  # Remove duplicate rows based on the Full name column
data["ชื่อ-นามสกุล ภาษาไทย (Full name in Thai)"] = data["ชื่อ-นามสกุล ภาษาไทย (Full name in Thai)"].str.replace("เเ", "แ")  # Replace 'เเ' with 'แ' character in a column
data.insert(data.columns.get_loc("เลขประจำตัวนิสิต (Student ID)"), "verifiedStudentID", '')  # Insert a new column for verified student ID
data.insert(data.columns.get_loc("ภาควิชาที่สังกัด (Major)"), "verifiedMajor", '')  # Insert a new column for verified major
data.to_excel("cleaneddata.xlsx", index=False)  # Write the cleaned data to a new Excel file

# Create a webdriver object with specified options
options = Options()
options.add_experimental_option("detach", True)  # Keep the browser open after the script execution
driver = webdriver.Chrome(options=options)

# Access the specified URL
driver.get("https://www2.reg.chula.ac.th/cu/general/PersonalInformation/InquiryNewStudentID/index.html")

# Find the input and output frames in the web page
inputframe, outputframe = driver.find_elements(By.XPATH, "//frame")

# Define a function to scrape data from the web page given a name and surname in Thai
def scrape(name, surname):
    # Switch to the default content and then switch to the input frame
    driver.switch_to.default_content()
    driver.switch_to.frame(inputframe)

    # Find the input fields for the name and surname in Thai
    nameThai = driver.find_element(By.NAME, "nameThai")
    surThai = driver.find_element(By.NAME, "surnameThai")
    submit = driver.find_element(By.NAME, "submit")

    # Enter the provided name and surname into the respective input fields
    nameThai.send_keys(name)
    surThai.send_keys(surname)

    # Click the submit button to initiate the search
    submit.click()

    # Clear the input fields for the next iteration
    nameThai.clear()
    surThai.clear()

    # Switch to the default content and then switch to the output frame
    driver.switch_to.default_content()
    driver.switch_to.frame(outputframe)

    # Find and extract the student ID and major information from the web page
    studentID = driver.find_element(By.XPATH, "/html/body/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table[1]/tbody/tr[3]/td[2]/p/b/font")
    major = driver.find_element(By.XPATH, "/html/body/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table[1]/tbody/tr[8]/td[2]/p/b/font")

    # Return the extracted student ID and major as a list
    return [studentID.text.replace(" ", ""), major.text]

# Define a function to scrape data from the web page given a name and surname in English
def scrapeeng(name, surname):
    # Switch to the default content and then switch to the input frame
    driver.switch_to.default_content()
    driver.switch_to.frame(inputframe)

    # Find the input fields for the name and surname in English
    nameENG = driver.find_element(By.NAME, "nameEng")
    surENG = driver.find_element(By.NAME, "surnameEng")
    submit = driver.find_element(By.NAME, "submit")

    # Enter the provided name and surname into the respective input fields
    nameENG.send_keys(name)
    surENG.send_keys(surname)

    # Click the submit button to initiate the search
    submit.click()

    # Clear the input fields for the next iteration
    nameENG.clear()
    surENG.clear()

    # Switch to the default content and then switch to the output frame
    driver.switch_to.default_content()
    driver.switch_to.frame(outputframe)

    # Find and extract the student ID and major information from the web page
    studentID = driver.find_element(By.XPATH, "/html/body/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table[1]/tbody/tr[3]/td[2]/p/b/font")
    major = driver.find_element(By.XPATH, "/html/body/table/tbody/tr/td[2]/table/tbody/tr/td[2]/table[1]/tbody/tr[8]/td[2]/p/b/font")

    # Return the extracted student ID and major as a list
    return [studentID.text.replace(" ", ""), major.text]

# Read data from the cleaned Excel file into a DataFrame
df = pd.read_excel("cleanedData.xlsx")

# Iterate over each row in the DataFrame and perform the scraping operations
for cell in df.index:
    try:
        # Extract the Thai name and perform scraping
        name = str(df.loc[cell, "ชื่อ-นามสกุล ภาษาไทย (Full name in Thai)"]).strip().split()
        Id, major = scrape(name[0], name[1])
        df.loc[cell, "verifiedStudentID"] = Id
        df.loc[cell, "verifiedMajor"] = major
    except Exception as e:
        try:
            # If an exception occurs, try extracting the English name and perform scraping
            nameeng = str(df.loc[cell, "ชื่อ-นามสกุล ภาษาอังกฤษ (Full name in English)"]).strip().split()
            Id, major = scrapeeng(nameeng[0], nameeng[1])
            df.loc[cell, "verifiedStudentID"] = Id
            df.loc[cell, "verifiedMajor"] = major
        except Exception as e:
            # If there's an error in both cases, print the row number and name for debugging purposes
            print(f"{cell} {name}")

# Close the webdriver
driver.close()

# Write the resulting DataFrame to a new Excel file
df.to_excel("resultData.xlsx", index=False)
