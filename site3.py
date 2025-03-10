import os
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from openpyxl import Workbook
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


options = Options()
# options.add_argument("--headless")  # Uncomment if you want to run the browser in headless mode

# Initialize WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Create an Excel workbook and add the headers
wb = Workbook()
ws = wb.active
ws.append(['Hotel Name', 'Address', 'Location', 'Email', 'Phone Number'])

# Loop through pages (from page 1 to page 55)
for page in range(1, 8):
    print(f"Scraping page {page}...")

    # Construct the URL with the current page number
    url = f'https://nidhi.tourism.gov.in/home/directory?categoryCode=09&category_name=Destinations%20and%20Attractions&pageno={page}'
    driver.get(url)

    # Wait for the listings to load (use WebDriverWait for better handling of dynamic content)
    try:
        WebDriverWait(driver, 3).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.listing-block'))
        )
    except Exception as e:
        print(f"Error waiting for page {page}: {e}")
        continue

    # Get all the listings on the current page
    listings = driver.find_elements(By.CSS_SELECTOR, '.listing-block')
    print(f"Found {len(listings)} listings on page {page}.")

    # Iterate through each listing and extract the required data
    for listing in listings:
        try:
            hotel_name = listing.find_element(By.CSS_SELECTOR, '.hotel-heading').text
            address = listing.find_element(By.CSS_SELECTOR, '.address').text
            location = listing.find_element(By.CSS_SELECTOR, '.share-location').text
            email = listing.find_element(By.CSS_SELECTOR, '.mail-details').text
            phone = listing.find_elements(By.CSS_SELECTOR, '.mail-details')[1].text

            print(f"Hotel Name: {hotel_name}")
            print(f"Address: {address}")
            print(f"Location: {location}")
            print(f"Email: {email}")
            print(f"Phone: {phone}")
            print("-" * 50)

            # Append the data to the Excel sheet
            ws.append([hotel_name, address, location, email, phone])
        except Exception as e:
            print(f"Error in listing on page {page}: {e}")

# Save the data to an Excel file
wb.save('hotels_data.xlsx')

# Close the driver
driver.quit()
