import pandas as pd
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# Function to read company names from an Excel file
def read_company_names_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Company Name'].tolist()


# Function to check if a URL is a LinkedIn profile
def is_linkedin_profile_url(url: str):
    return url.startswith('https://in.linkedin.com/in')
# def is_linkedin_profile_url(url: str):
#     # Ensure the URL starts with the LinkedIn base URL for profiles
#     # and explicitly exclude company URLs
#     return url.startswith('https://www.linkedin.com/in') and '/company/' not in url


# Function to search and scrape LinkedIn URLs
def search_and_scrape(company_name):
    url = 'https://www.google.com/'
    browser = webdriver.Chrome()
    browser.get(url)
    textarea = browser.find_element(By.TAG_NAME, "textarea")
    search_query = f'{company_name} Linkedin Profile' # TEXT TO SEARCH
    textarea.send_keys(search_query)
    textarea.send_keys(Keys.ENTER)
    WebDriverWait(browser, 10000).until(EC.presence_of_element_located((By.ID, 'logo')))
    page_source = browser.page_source
    soup = BeautifulSoup(page_source, 'lxml')
    elements = soup.find_all('a', attrs={'jsname': 'UWckNb'})

    links = []
    for link in elements:
        if len(links) >= 1:  # Limit to first two links
            break
        try:
            href = link['href']
            if is_linkedin_profile_url(href):
                links.append(href)
        except KeyError:
            continue
    browser.quit()  # Ensure the browser is closed after the search
    return links


# Function to append URLs to an Excel file
# def append_urls_to_excel(file_path, company_name, urls):
#     # Try to read the existing Excel file
#     try:
#         df_existing = pd.read_excel(file_path)
#     except FileNotFoundError:
#         # If the file does not exist, create a new DataFrame
#         df_existing = pd.DataFrame(columns=['Company Name', 'LinkedIn URL'])
#
#     # Create a new DataFrame with the new data
#     df_new = pd.DataFrame({'Company Name': [company_name]*len(urls), 'LinkedIn URL': urls})
#
#     # Append the new data to the existing DataFrame
#     df_final = pd.concat([df_existing, df_new], ignore_index=True)
#
#     # Write the combined DataFrame back to the Excel file
#     df_final.to_excel(file_path, index=False)


def append_urls_to_excel(file_path, company_name, urls):
    # Ensure the URLs list has at least one entry, even if it's "N/A"
    if not urls:
        urls = ["N/A"]
    else:
        # Ensure all URLs are treated as strings
        urls = [str(url) for url in urls]

    try:
        # Attempt to read the existing Excel file into a DataFrame
        df_existing = pd.read_excel(file_path, dtype={'Company Name': str, 'LinkedIn URL': str})
    except FileNotFoundError:
        # Initialize a DataFrame with appropriate columns if the file doesn't exist
        df_existing = pd.DataFrame(columns=['Company Name', 'LinkedIn URL'])

    # Check if the company name already exists in the DataFrame
    if company_name in df_existing['Company Name'].values:
        # Find the row index where the company name matches and update the LinkedIn URL
        for row_index in df_existing.index[df_existing['Company Name'] == company_name].tolist():
            df_existing.at[row_index, 'LinkedIn URL'] = ', '.join(urls)  # Joining URLs with a comma if there are multiple
    else:
        # Append new rows for the company with each URL
        for url in urls:
            df_existing = df_existing.append({'Company Name': company_name, 'LinkedIn URL': url}, ignore_index=True)

    # Write the DataFrame back to the Excel file, ensuring string format for URLs
    df_existing.to_excel(file_path, index=False)



# Main code
file_path = 'companies.xlsx'
company_names = read_company_names_from_excel(file_path)
for company_name in company_names:
    urls = search_and_scrape(company_name)
    if urls:  # Check if any URL was found
        append_urls_to_excel(file_path, company_name, urls)
