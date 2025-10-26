import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import os


# folder for files
base_path = './excel_files'


# input and output files
input_file = os.path.join(base_path, 'companies.xlsx')
output_file = os.path.join(base_path, 'companies_with_tax.xlsx')

# read company list
df = pd.read_excel(input_file)

# store all results
results = []

# user agent header
headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/127.0.0.0 Safari/537.36"
}


# go through each company code
for code in df["Code"]:
    url = f"https://ariregister.rik.ee/eng/company/{code}"
    print(f"Processing {code} ...")

    try:
        r = requests.get(url, headers=headers, timeout=5)
        if r.status_code != 200:
            raise Exception(f"HTTP {r.status_code}")
    except Exception as e:
        print(f"Error loading {code}: {e}")
        results.append({
            "Code": code,
            "Period": None,
            "State taxes": None,
            "Taxes on workforce": None,
            "Taxable turnover": None,
            "Number of employees": None
        })
        time.sleep(3)
        continue

    soup = BeautifulSoup(r.text, "html.parser")
    tax_info = soup.find("div", {"id": "tax-information"})

    
    if not tax_info:
        print(f"No tax data for {code}")
        results.append({
            "Code": code,
            "Period": None,
            "State taxes": None,
            "Taxes on workforce": None,
            "Taxable turnover": None,
            "Number of employees": None
        })
        time.sleep(3)
        continue

    text_block = tax_info.get_text(" ", strip=True)

    # extract reporting period
    period = None
    for line in text_block.split("\n"):
        if "Taxes paid" in line:
            period = line.strip()

    # extract value by label
    def extract_value(label):
        el = tax_info.find(string=lambda t: label in t)
        if el and el.parent.find_next("td"):
            return el.parent.find_next("td").get_text(strip=True)
        return None

    
    results.append({
        "Code": code,
        "Period": period,
        "State taxes": extract_value("State taxes"),
        "Taxes on workforce": extract_value("Taxes on workforce"),
        "Taxable turnover": extract_value("Taxable turnover"),
        "Number of employees": extract_value("Number of employees"),
    })

    # wait before next request
    time.sleep(4)


# save results
out_df = pd.DataFrame(results)
out_df.to_excel(output_file, index=False)


print()
print("Results saved successfully")
print(f"File: {output_file}")
print("Done")
