# CorpScrapper

**CorpScrapper** is a web scraping tool designed to extract detailed information about corporations/companies. Originally developed for a potential project, this tool takes company names from an Excel file, searches specified websites, and extracts data such as addresses, zip codes, presidents' names, and mailing addresses. It's a modular, user-friendly script that provides terminal prompts and progress updates for seamless navigation.

---

## Features

- **Excel Integration**: Reads company names from an input Excel file.
- **Customizable Targets**: Placeholder URL allows easy modification for other target websites.
- **Data Extraction**:
  - Company address, zip code, president's name, mailing address, and more.
  - Missing data automatically marked as "N/A."
- **Multi-Step Navigation**: Automatically follows redirects and navigates through multiple pages to locate and extract the required information.
- **Prompted Workflow**: Guided prompts at every step, displaying extracted information in real-time.
- **Progress Tracking**:
  - Success message with elapsed and estimated time for task completion.
  - Frequency of data extraction is customizable.
- **Output**:
  - Results saved in an organized format for easy analysis.

![Image-1](https://github.com/user-attachments/assets/083ddbe7-6464-4ac5-bf80-754ef9b1c47c)
![Image-2](https://github.com/user-attachments/assets/adb3c7a7-69f3-4928-886b-fc4fbc2d2b4a)
![Image-3](https://github.com/user-attachments/assets/cfc82f23-be87-49a9-90e8-cd591685d91a)
![Image-4](https://github.com/user-attachments/assets/6851a83a-b203-4909-83d8-e6fbcdbac556)
![Image-5](https://github.com/user-attachments/assets/7aa9d490-c6c5-419e-a1cf-9236e86c9145)

---

## Getting Started

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/CorpScrapper.git

git clone https://github.com/mamq99/CorpScrapper.git

2. Navigate to the project folder:
   ```bash
   cd CorpScrapper
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
## Usage
1. Place your input Excel file in the project directory.
2. Open the script in a code editor to update:
  - The placeholder target URL.
  - Any desired data limits or frequency of extraction.
3. Run the script:
    ```bash
    python CorpScrapper.py
4. Follow the terminal prompts as the script processes each company.

## Output
- Extracted data is saved in a results file (e.g., Excel) in the output directory.
- Example:
  ```bash
  Company Name, Address, Zip Code, President Name, Mailing Address, Permanent Address

## Technologies Used
- Python: Core programming language.
- BeautifulSoup4 (BS4): For parsing website HTML.
- Selenium: For automating browser interactions.
- Pandas: For handling Excel files and data manipulation.

## Customization
- Update the placeholder URL to scrape other websites.
- Adjust frequency or data limits within the script to suit your needs.

## Limitations
- Missing information is marked as "N/A."
- The tool relies on the structure of target websites. Significant website changes may require updates to the script.

## Privacy Notice
- Original Excel and results files were removed for privacy reasons.

## License
This project is open-source with no license restrictions. Feel free to use, modify, or distribute it.  
