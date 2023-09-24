import requests
from bs4 import BeautifulSoup
import openpyxl


# Function to scrape data for a single company
def scrape_company_data(company_url):
    response = requests.get(company_url)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        
        # Extract data from the company's page
        company_name = soup.find("h1").text.strip()
        
		# Extracting number of employees
        employees = ""
        number_of_employees_label = soup.find('p', string='Number of employees')
        if number_of_employees_label:
          number_of_employees = number_of_employees_label.find_next('p').text.strip()
          employees = number_of_employees
        else:
          employees = "-"
          
		# Extracting founded date
        founded_on = ""
        founded_text = soup.find('p', string='Founded')
        if founded_text:
          founded_on = founded_text.find_next('p').text.strip().split(',', 1)[0]
        else:
          founded_on = "-"
          
		# Extracting funded amount 
        funded_amount = ""
        funded_amount_text = soup.find('span', string='Total Funding')
        if funded_amount_text:
          funded_amount = funded_amount_text.find_next('span').text.strip()
        else:
          funded_amount = "-"
          
        return {
            "Company Name": company_name,
            "Founded": founded_on,
            "Number of employees": employees,
            "Total Funding": funded_amount,
        }
    else:
        print(f"Failed to retrieve data for {company_url}")
        return None

# URL of the main page
main_url = "https://big-picture.com/fintech/germany.html"

# Send a GET request to the main page
response = requests.get(main_url)

if response.status_code == 200:
    soup = BeautifulSoup(response.text, "html.parser")
    
    # Find links to company pages
    company_links = [a['href'] for div in soup.find_all('div', class_='fintech-list-companies') for a in div.find_all('a')]
    # Print the extracted links
      
    # Initialize a list to store scraped data for each company
    company_data_list = []
    
    # Loop through each company link and scrape data
    for company_link in company_links:
        company_url = f"https://big-picture.com/fintech/{company_link}"
        company_data = scrape_company_data(company_url)
        if company_data:
            company_data_list.append(company_data)
    
    # Create a new Excel workbook and add a worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    worksheet.title = "Company Data"
    
    # Write the header row to the Excel sheet
    header = [
        "Company Name",
        "Founded",
        "Number of employees",
        "Total Funding"
    ]
    worksheet.append(header)
    
    # Write data for each company to the Excel sheet
    for company_data in company_data_list:
        row = [
            company_data["Company Name"],
            company_data["Founded"],
            company_data["Number of employees"],
            company_data["Total Funding"],
        ]
        worksheet.append(row)
    
    # Save the Excel workbook
    workbook.save("company_data.xlsx")
    print("Data saved to company_data.xlsx")
else:
    print(f"Failed to retrieve data from {main_url}")