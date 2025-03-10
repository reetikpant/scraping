from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd
import time

options = webdriver.ChromeOptions()
options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

url = 'https://indmount.org/IMF/contentDisplayTour'
driver.get(url)

time.sleep(5)

html = driver.page_source

soup = BeautifulSoup(html, 'html.parser')

data = []

panels = soup.find_all('div', class_='panel panel-info')

if panels:
    print(f"Found {len(panels)} panels")
    for panel in panels:
        name_tag = panel.find('span', style='font-weight: bold; font-size: large')
        if name_tag:
            name = name_tag.text.strip()
        else:
            name_tag_alt = panel.find('div', class_='panel-heading').find('span')
            name = name_tag_alt.text.strip() if name_tag_alt else 'N/A'

        email_tag = panel.find('div', class_='pull-right')
        email = email_tag.text.replace('E-Mail:', '').strip() if email_tag else 'N/A'

        address_tag = panel.find('div', class_='panel-footer').find('div', class_='col-sm-8').find('b')
        address = address_tag.text.strip() if address_tag else 'N/A'

        phone_tag = panel.find('div', class_='pull-right col-sm-4')
        phone = phone_tag.text.replace('Phone:', '').strip() if phone_tag else 'N/A'

        phone_numbers = phone.split(',') if phone != 'N/A' else []
        phone_1 = phone_numbers[0].strip() if len(phone_numbers) > 0 else 'N/A'
        phone_2 = phone_numbers[1].strip() if len(phone_numbers) > 1 else 'N/A'

        data.append({
            'Name': name,
            'Email': email,
            'Address': address,
            'Phone 1': phone_1,
            'Phone 2': phone_2
        })

    df = pd.DataFrame(data)
    df.to_excel('scraped_data.xlsx', index=False)

    print('Data has been successfully scraped and saved to scraped_data.xlsx')
else:
    print("No panels found on the page.")

driver.quit()
