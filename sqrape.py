import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook

def search_and_extract_contact_info(url, wb, ws):
    # Send a GET request to the URL
    response = requests.get(url)
    
    # Check if the request was successful
    if response.status_code == 200:
        # Parse the HTML content
        soup = BeautifulSoup(response.content, 'html.parser')
        
        # Find all text content
        text_content = soup.get_text()
        
        # Compile regular expression patterns for email, mobile, and contact
        email_pattern = re.compile(r'email\s*:\s*([\w.-]+@[\w.-]+)', re.IGNORECASE)
        mobile_pattern = re.compile(r'mobile\s*:\s*([\d\s-]+)', re.IGNORECASE)
        contact_pattern = re.compile(r'contact\s*:\s*([\w\d\s-]+)', re.IGNORECASE)
        
        # Search for email, mobile, and contact information
        email_match = email_pattern.search(text_content)
        mobile_match = mobile_pattern.search(text_content)
        contact_match = contact_pattern.search(text_content)
        
        # Extract and compile contact information
        website = url
        email = email_match.group(1) if email_match else ''
        mobile = mobile_match.group(1) if mobile_match else ''
        contact = contact_match.group(1) if contact_match else ''
        
        # Append data to Excel sheet
        ws.append([website, email, mobile, contact])
        
        print(f"Contact information extracted for {url}")
    else:
        # Print an error message if the request failed
        print(f"Failed to fetch data from {url}. Status code: {response.status_code}")

# Create Excel workbook and sheet
wb = Workbook()
ws = wb.active

# Add headers to Excel sheet
ws.append(['Website', 'Email', 'Mobile', 'Contact'])

# List of websites to search
websites = [
    'https://www.lavillaaphro.com/contact-la-villa-aphro/',
    'https://erata-hotel.business.site/',
    'https://www.businessghana.com/site/directory/2-star-hotels/453853/Rock-City-Hotel',
    'https://www.ghanayello.com/company/46466/Rock_City_Hotel',
    'https://www.businessghana.com/site/directory/2-star-hotels/438247/Prestige-Suites-Hotel-Gh',
    'https://www.ghanayello.com/company/46466/Rock_City_Hotel',
    'https://ange-hill-hotel.accra-hotels-gh.com/en/',
    'https://www.palomahotel.com/',
    'https://the-pearl-in-the-city.business.site/',
    # Add more websites here
]

# Iterate over websites and search for contact information
for website in websites:
    search_and_extract_contact_info(website, wb, ws)

# Save Excel file
wb.save('contact_info.xlsx')
print("Contact information extracted and saved to contact_info.xlsx")
