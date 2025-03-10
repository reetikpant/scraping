from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import platform
import os

def get_google_ads_data_selenium(domain, start_date, end_date, region="IN", platform_name="YOUTUBE"):
    """
    Fetches data from the Google Ads Transparency Center using Selenium and ChromeDriverManager.
    """
    url = f"https://adstransparency.google.com/?region={region}&domain={domain}&platform={platform_name}&start-date={start_date}&end-date={end_date}"

    try:
        # Set up Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920x1080")

        # Check if running on Mac ARM64
        if platform.system() == "Darwin" and platform.machine() == "arm64":
            # Manually specify ChromeDriver path for Mac ARM64
            chromedriver_path = "/usr/local/bin/chromedriver"  # Common location
            if not os.path.exists(chromedriver_path):
                print("ChromeDriver not found. Please install via Homebrew: brew install chromedriver")
                return []
            service = Service(chromedriver_path)
        else:
            # Use WebDriverManager for other systems
            service = Service(ChromeDriverManager().install())

        driver = webdriver.Chrome(service=service, options=chrome_options)

        driver.get(url)
        time.sleep(10)  # Wait for page to load

        # Find creative previews
        try:
            creative_previews = driver.find_elements(By.CSS_SELECTOR, "creative-preview")
        except Exception as e:
            print(f"Error finding creative previews: {e}")
            driver.quit()
            return []

        ad_data = []
        for preview in creative_previews:
            try:
                a_tag = preview.find_element(By.CSS_SELECTOR, "a")
                ad_url = a_tag.get_attribute("href")
                
                advertiser_name_element = preview.find_element(By.CSS_SELECTOR, ".advertiser-name")
                advertiser_name = advertiser_name_element.text
                
                verified_element = preview.find_element(By.CSS_SELECTOR, ".verified")
                verified_status = verified_element.text

                ad_data.append({
                    "url": ad_url,
                    "advertiser_name": advertiser_name,
                    "status": verified_status
                })
            except Exception as e:
                print(f"Error extracting data from a creative preview: {e}")

        driver.quit()
        return ad_data

    except Exception as e:
        print(f"Error during Selenium execution: {e}")
        return []

# Example Usage
if __name__ == "__main__":
    domain_name = "atlys.com"
    start_date = "2025-01-31"
    end_date = "2025-02-10"

    ads = get_google_ads_data_selenium(domain_name, start_date, end_date)

    if ads:
        for ad in ads:
            print(f"URL: {ad['url']}")
            print(f"Advertiser Name: {ad['advertiser_name']}")
            print(f"Status: {ad['status']}")
            print("-" * 20)
    else:
        print("Failed to retrieve ad data.")
