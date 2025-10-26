# Company Tax Data Scraper

This project takes company codes from an Excel file, sends web requests to the Estonian Business Register website, and extracts tax and workforce data.  
It saves all results to a new Excel file for further analysis.



## Features

- Reads company codes from Excel
- Sends web requests to fetch company tax data
- Extracts tax period, state taxes, turnover, and employees
- Waits between requests to avoid server overload
- Handles missing or invalid data gracefully
- Saves a clean Excel report



## Requirements

- Python 3.9 or newer  
- Libraries: `requests`, `beautifulsoup4`, `pandas`, `openpyxl`

Install them with:

```bash
pip install requests beautifulsoup4 pandas openpyxl
