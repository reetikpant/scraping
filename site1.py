import requests
from bs4 import BeautifulSoup
import pandas as pd

# Predefined column names
columns = [
    "Organization Name", "Reg. No", "Vat No", "Address",
    "Country", "Website URL", "Email", "Telephone number",
    "Mobile number", "Fax", "PO Box", "Key Person", "Establishment Date"
]

# Base URL for scraping
base_url = "https://www.taan.org.np/members/"

# Initialize a list to store the data
all_data = []

# Loop through member IDs from 1 to 10
for member_id in range(1, 3001):
    url = f"{base_url}{member_id}"
    try:
        # Fetch the page content
        response = requests.get(url)
        response.raise_for_status()  # Raise error for HTTP issues

        # Check if the URL redirects to 404
        if response.url == "https://www.taan.org.np/404":
            print(f"URL {url} leads to a 404 page. Skipping.")
            continue

        html_content = response.text
        soup = BeautifulSoup(html_content, "html.parser")

        # Find the ul with the class 'list-group'
        list_group = soup.find("ul", class_="list-group small")

        # Extract the text of each list-group-item and clean the text
        data = []
        if list_group:
            items = list_group.find_all("li", class_="list-group-item")
            for item in items:
                # Remove strong tags and clean the text
                for strong in item.find_all("strong"):
                    strong.decompose()
                text = item.get_text(strip=True)
                # Remove the ':' and append to data
                text = text.replace(":", "").strip()
                data.append(text)

        # Map data to column names
        record = {columns[i]: data[i] if i < len(data) else "" for i in range(len(columns))}
        record["URL"] = url  # Add the URL for reference
        all_data.append(record)

    except requests.exceptions.RequestException as e:
        print(f"Error fetching the URL {url}: {e}")

# Save the data to an Excel file
if all_data:
    df = pd.DataFrame(all_data)
    output_file = "taan_members_data.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Data saved to {output_file}")
else:
    print("No data to save.")
