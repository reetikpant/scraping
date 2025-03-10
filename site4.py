from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd
import time

# Function to scrape the webpage and extract the necessary data
def scrape_and_save_to_excel_with_selenium(base_url, start_page, end_page, excel_filename):
    # Set up Chrome options for headless operation (no GUI)
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")  # Uncomment to run in headless mode
    
    # Initialize Chrome WebDriver using webdriver-manager to auto-download the correct driver version
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    # List to store extracted data
    all_data = []

    # Loop through all pages from start_page to end_page
    for page_num in range(start_page, end_page + 1):
        url = f"{base_url}{page_num}"  # Create the URL for the current page
        print(f"Scraping page {page_num}: {url}")
        
        # Open the URL in the browser
        driver.get(url)

        # Wait for the page to load completely
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "a.item.hovercard")))

        # Allow some additional time for content to load
        time.sleep(3)

        # Get the page source after JavaScript has executed
        page_source = driver.page_source

        # Parse the page source with BeautifulSoup
        soup = BeautifulSoup(page_source, 'html.parser')

        # Extract the relevant information (image URL, title, and author)
        data = []
        
        # Find all the <a> tags with class 'item hovercard'
        a_tags = soup.find_all('a', class_='item hovercard')
        if a_tags:  # Check if any 'a' tags were found
            for a_tag in a_tags:  # Loop through each <a> tag
                # Get the img tag that is a direct child of the 'a' tag
                img_tag = a_tag.find('img', recursive=False)
                img_url = img_tag['src'] if img_tag else None
                
                title_tag = a_tag.find('h5', class_='text-overflow title-hover-content')
                title = title_tag.text.strip() if title_tag else None

                author_tag = a_tag.find('h5', class_='text-overflow author-label mg-bottom-xs')
                author = author_tag.find('em').text.strip() if author_tag else None

                view_tag = a_tag.find('span', class_='sub-hover')
                view = None
                if view_tag:
                    eye_span = view_tag.find('i', class_='fa fa-eye myicon-right')
                    if eye_span:
                        view = eye_span.find_next_sibling(string=True).strip()

                # Append the extracted information to the data list
                data.append([img_url, title, author, view])
            
        else:
            print(f"No items found for page {page_num}")

        # Add the data from the current page to the overall list
        all_data.extend(data)

    # Close the driver after scraping
    driver.quit()

    # Create a DataFrame from the extracted data
    df = pd.DataFrame(all_data, columns=['Image URL', 'Title', 'Author', 'View'])
    print(df)

    # Save the DataFrame to an Excel file
    df.to_excel(excel_filename, index=False, engine='openpyxl')
    print(f"Data has been successfully saved to {excel_filename}")

# URL of the webpage to scrape (base URL with pagination)
base_url = 'https://photo.ntb.gov.np/ajax/latest?page='

# Output Excel filename
excel_filename = 'scraped_data_with_selenium.xlsx'

# Loop over pages 1 to 93
scrape_and_save_to_excel_with_selenium(base_url, start_page=1, end_page=93, excel_filename=excel_filename)
