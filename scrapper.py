from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from bs4 import BeautifulSoup
from time import sleep
import os
import time  # Ensure this is imported at the top of your file


# If you wish to log then uncomment the below logging code

# import logging
# from concurrent.futures import ThreadPoolExecutor

# Configure logging
# logging.basicConfig(
#     filename='scraper.log',  # Log file name
#     level=logging.INFO,  # Set the logging level
#     format='%(asctime)s - %(levelname)s - %(message)s'  # Log format
# )

class SimpleScraper:
    def __init__(self):
        # Setup Chrome with additional options
        options = Options()
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--headless')  # Enable headless mode
        # options.add_argument('--start-maximized')  # Start with maximized window
        
        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
        
        # Set implicit wait time
        self.driver.implicitly_wait(10) 
        
        self.wait = WebDriverWait(self.driver, 20)
        
    def search_company(self, company_name):
        """Search for a company on sunbiz.org"""
        try:
            # Use the complete URL
            print(f"Accessing website for: {company_name}")
            self.driver.get("Your Destination Website's URL Comes here") 
            sleep(3)  # Wait for page to load
            
            # Print current URL for debugging
            print(f"Current URL: {self.driver.current_url}")
            
            # Search box interaction
            search_box = self.wait.until(
                EC.presence_of_element_located((By.ID, "SearchTerm"))
            )
            search_box.clear()
            search_box.send_keys(company_name)
            
            search_button = self.driver.find_element(By.XPATH, "/html/body/div[1]/div[1]/div[2]/div/div/form/div[2]/div[2]/input")
            search_button.click()
            # sleep(6)  # Increased wait time for results to load

            #Wait for the page to load and then load the table of information and select the company
            company_table_body = self.wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '#search-results table tbody')),
                print("found the table")
            )
            sleep(2)  # Increased wait time
            company_table_body.find_element(By.CSS_SELECTOR, 'tr td:nth-child(1) a').click()
            print("found the row and clicked")
            sleep(4)  # Increased wait time
            return True
            
        except Exception as e:
            print(f"Error searching for {company_name}: {str(e)}")
            return False

    def get_company_details(self):
        """Extract company details from the page"""
        try:
            print("Extracting company details...")
            details = {}
            
            # Function to extract address details
            def extract_address_details(address_text, extract_city_state_zip=False):
                address_lines = address_text.splitlines()  # Split into lines
                address_lines = [line.strip() for line in address_lines if line.strip()]  # Remove empty lines

                last_address_line = ""
                if address_lines:
                    last_address_line = address_lines[-1]  # Get the last line for city, state, zip

                # Join address lines into a single string
                full_address = " ".join(address_lines)

                # Initialize city, state, and zip code
                city, state, zip_code = "", "", ""
                if extract_city_state_zip and last_address_line:
                    state_zip_parts = last_address_line.split(' ')
                    if len(state_zip_parts) >= 2:
                        city = ' '.join(state_zip_parts[:-2])  # Join all but the last two parts as city
                        state = state_zip_parts[-2]  # Second last part as state
                        zip_code = state_zip_parts[-1]  # Last part as zip code

                return full_address, city, state, zip_code

            # Principal Address
            try:
                principal_address = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.detailSection:nth-child(4) span:nth-child(2)'))
                ).text
                print("Found principal address")
                print(f"{principal_address}")
                details['principal_address'], _, _, _ = extract_address_details(principal_address)  # Only get full address

            except Exception as e:
                details['principal_address'] = "Not found"
                print(f"Principal address not found: {str(e)}")
            
            # Mailing Address
            try:
                mailing_address = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.detailSection:nth-child(5) span:nth-child(2)'))
                ).text
                print("Found mailing address")
                print(f"{mailing_address}")
                details['mailing_address'], details['mailing_city'], details['mailing_state'], details['mailing_zip_code'] = extract_address_details(mailing_address, extract_city_state_zip=True)  # Get full address and city/state/zip

            except Exception as e:  
                details['mailing_address'] = "Not found"
                details['mailing_city'] = details['mailing_state'] = details['mailing_zip_code'] = "Not found"
                print(f"Mailing address not found: {str(e)}")

            # President Information
            try:
                officer_details = self.wait.until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, '.detailSection:nth-child(7)'))
                )
                officer_details_content = officer_details.get_attribute('outerHTML')

                soup = BeautifulSoup(officer_details_content, 'html.parser')

                # Get all text from the officer details section
                #################### VERY IMPORTANT ADDITION TO THE CODE ################ MADE IT LESS RESOURCE-HEAVY
                all_text = soup.get_text(separator="\n")  # Use separator to maintain line breaks
                lines = all_text.splitlines()  # Split into lines

                # Remove empty lines
                lines = [line.strip() for line in lines if line.strip()]  # Filter out empty lines

                president_name = "Not found"
                president_address = "Not found"

                # Initialize a flag to track if we found the president title
                found_president = False

                # Iterate through the lines to find the title
                for i in range(len(lines)):
                    line = lines[i].strip()  # Clean up whitespace
                    
                    # Check if the current line contains a title for President
                    if 'Title' in line and any(pos in line for pos in ['President', 'Pres', 'Title P']):
                        found_president = True
                        print(f"President found? {found_president}")

                        # The name is in the next line
                        if i + 1 < len(lines):
                            president_name = lines[i + 1].strip()  # Get the next line for the name

                        details['president_name'] = president_name
                        print(f"President Name: {president_name}")

                        # The address is usually in the following lines
                        if i + 2 < len(lines):
                            # Collect address lines until the next title or empty line
                            address_lines = []
                            last_address_line = ""  # Variable to hold the last address line
                            for j in range(i + 2, len(lines)):
                                if lines[j].strip() == "" or 'Title' in lines[j]:  # Stop if we hit an empty line or another title
                                    break
                                address_lines.append(lines[j].strip())
                                last_address_line = lines[j].strip()  # Update the last address line

                            # Join address lines into a single string
                            president_address = " ".join(address_lines)  
                            details['president_address'] = president_address
                            print(f"President Address: '{president_address}'")

                            # Now extract city, state, and zip code from the last address line
                            if last_address_line:
                                print(f"Last Address Line for City, State, Zip: '{last_address_line}'")  # Debugging line
                                state_zip_parts = last_address_line.split(' ')
                                if len(state_zip_parts) >= 2:
                                    details['president_city'] = ' '.join(state_zip_parts[:-2])  # Join all but the last two parts as city
                                    details['president_state'] = state_zip_parts[-2]  # Second last part as state
                                    details['president_zip_code'] = state_zip_parts[-1]  # Last part as zip code
                                else:
                                    details['president_city'] = details['president_state'] = details['president_zip_code'] = ''
                            else:
                                details['president_city'] = details['president_state'] = details['president_zip_code'] = ''

                        break  # Exit the loop after finding the president details

                if not found_president:
                    print(f"President found? {found_president}")
                    details['president_name'] = "N/A"
                    details['president_address'] = "N/A"
                    details['president_city'] = details['president_state'] = details['president_zip_code'] = "N/A"

            except Exception as e:
                print(f"Exception: President Detail not found: {str(e)}")
                details['president_name'] = "N/A"
                details['president_address'] = "N/A"
                details['president_city'] = details['president_state'] = details['president_zip_code'] = "N/A"

            return details
            
        except Exception as e:
            print(f"Error getting details: {str(e)}")
            return {}

    def process_companies(self, excel_file):
        """Process list of companies from Excel file"""
        try:
            # Read Excel file with explicit path
            print(f"Reading Excel file: {excel_file}")
            df = pd.read_excel(excel_file)
            
            # Print columns for debugging
            print("\nAvailable columns in Excel file:")
            print(df.columns.tolist())

            #You can adjust this below here for your own needs
            companies = df.iloc[1:101, 0]  # Adjusted to process the first 100 companies
            
            results = []
            start_time = time.time()  # Start timer
            
            for i, company in enumerate(companies, 1):
                print(f"\nProcessing company {i}/100: {company}")
                
                if self.search_company(company):
                    # Wait for the company details to load
                    sleep(6)  # Increase wait time if necessary
                    details = self.get_company_details()
                    if details:
                        # Extract city, state, and zip from mailing address

                        results.append({
                            'Company Name': company,
                            'Principal Address': details.get('principal_address', ''),
                            'Mailing Address': details.get('mailing_address', ''),
                            'City': details.get('mailing_city', ''),
                            'State': details.get('mailing_state', ''),
                            'Zip': details.get('mailing_zip_code', ''),
                            'President Name': details.get('president_name', ''),
                            'President Address': details.get('president_address', ''),
                            'President City': details.get('president_city', ''),
                            'President State': details.get('president_state', ''),
                            'President Zip': details.get('president_zip_code', '')
                        })
                        print(f"Successfully processed: {company}")
                
                # Added a timer to get guage the elapsed and estimated remaining time for the scrapping 
                # Calculate elapsed time
                elapsed_time = time.time() - start_time
                estimated_time_per_company = elapsed_time / i  # Average time per company
                remaining_companies = len(companies) - i
                estimated_total_time = estimated_time_per_company * len(companies)
                estimated_remaining_time = estimated_total_time - elapsed_time
                
                # Convert estimated remaining time to minutes and seconds
                minutes, seconds = divmod(estimated_remaining_time, 60)
                
                # Print estimated time remaining in minutes and seconds
                print(f"Estimated time remaining: {int(minutes)} minutes and {int(seconds)} seconds")
                
                sleep(2)  # Consider adjusting this as well
                
            return pd.DataFrame(results)
            
        except Exception as e:
            print(f"Error processing companies: {str(e)}")
            return pd.DataFrame()

    def close(self):
        """Close the browser"""
        self.driver.quit()

def main():
    scraper = SimpleScraper()
    try:
        # Use the full path for the Excel file
        results_file = 'results.xlsx'
        
        # Create an empty DataFrame to hold results if the file does not exist
        if not os.path.exists(results_file):
            results = pd.DataFrame(columns=['Company Name', 'Principal Address', 'Mailing Address', 'City', 'State', 'Zip', 'President Name', 'President Address', 'President City', 'President State', 'President Zip'])
        else:
            results = pd.read_excel(results_file)

        # Process companies
        new_results = scraper.process_companies(r'D:\\CorpScrapper\data\\Your_file_name_here')
        
        # Append new results to the existing DataFrame
        results = pd.concat([results, new_results], ignore_index=True)
        
        # Save the updated DataFrame back to the Excel file
        results.to_excel(results_file, index=False)
        print("\nResults saved to results.xlsx")
        
    except Exception as e:
        print(f"Error in main execution: {str(e)}")
    finally:
        scraper.close()

if __name__ == "__main__":
    main()